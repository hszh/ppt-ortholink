using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OrthoLink
{
    /// <summary>
    /// Raw clipboard access. PowerPoint can only bring a custom-geometry shape onto a slide through the
    /// clipboard, so the add-in borrows it: back up every format byte for byte, put our shape on it, paste,
    /// put everything back. Both what we put there and the restored copy are marked so that clipboard
    /// managers (Ditto and the like) and Windows clipboard history do not record them.
    /// </summary>
    internal static class Clip
    {
        public const string Gvml = "Art::GVML ClipFormat";          // how Office moves shapes between programs
        private const string IgnoreFormat = "Clipboard Viewer Ignore";
        private const string ExcludeFormat = "ExcludeClipboardContentFromMonitorProcessing";

        private const uint CF_BITMAP = 2, CF_METAFILEPICT = 3, CF_PALETTE = 9, CF_ENHMETAFILE = 14;
        private const uint CF_OWNERDISPLAY = 0x80, CF_DSPBITMAP = 0x82, CF_DSPMETAFILEPICT = 0x83, CF_DSPENHMETAFILE = 0x8E;
        private const uint CF_PRIVATEFIRST = 0x200, CF_GDIOBJLAST = 0x3FF;
        private const uint GMEM_MOVEABLE = 0x0002;
        private const uint QS_SENDMESSAGE = 0x0040, PM_NOREMOVE = 0x0000, PM_QS_SENDMESSAGE = QS_SENDMESSAGE << 16;

        private static readonly long MaxBytes = Environment.Is64BitProcess ? 256L << 20 : 64L << 20;
        private static NativeWindow _owner;

        /// <summary>What was on the clipboard, in the order the program that put it there offered it.</summary>
        public sealed class Snapshot
        {
            internal readonly List<KeyValuePair<uint, byte[]>> Items = new List<KeyValuePair<uint, byte[]>>();
        }

        public static Snapshot Save()
        {
            var snap = new Snapshot();
            Open(2000);
            try
            {
                long budget = MaxBytes;
                for (uint f = EnumClipboardFormats(0); f != 0; f = EnumClipboardFormats(f))
                {
                    if (Skip(f)) continue;
                    byte[] data = Read(f, ref budget);
                    if (budget < 0)
                    {
                        Log.Write("clipboard holds more than " + (MaxBytes >> 20) + " MB, not backed up; it will be left empty");
                        snap.Items.Clear();
                        break;
                    }
                    if (data != null) snap.Items.Add(new KeyValuePair<uint, byte[]>(f, data));
                }
            }
            finally { CloseClipboard(); }
            return snap;
        }

        /// <summary>Replace the clipboard content by a single format.</summary>
        public static void Put(string format, byte[] data)
        {
            Open(2000);
            try
            {
                EmptyClipboard();
                Set(RegisterClipboardFormat(format), data);
                MarkPrivate();
            }
            finally { CloseClipboard(); }
        }

        /// <summary>Put a snapshot back. An empty snapshot leaves the clipboard empty.</summary>
        public static void Restore(Snapshot snap)
        {
            Open(5000);
            try
            {
                EmptyClipboard();
                if (snap == null || snap.Items.Count == 0) return;
                foreach (var it in snap.Items) Set(it.Key, it.Value);
                MarkPrivate();   // it went into the history already when it was copied
            }
            finally { CloseClipboard(); }
        }

        public static void Shutdown()
        {
            try { if (_owner != null) _owner.DestroyHandle(); } catch { }
            _owner = null;
        }

        // ------------------------------------------------------------------ internals

        private static void Open(int timeoutMs)
        {
            if (_owner == null)
            {
                var w = new NativeWindow();
                w.CreateHandle(new CreateParams { Caption = "OrthoLink clipboard", Parent = new IntPtr(-3) });   // message-only window
                _owner = w;
            }
            var until = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (!OpenClipboard(_owner.Handle))
            {
                if (DateTime.UtcNow > until)
                {
                    string holder = Holder();
                    Log.Write("clipboard kept open by " + holder + " for more than " + timeoutMs + " ms");
                    throw new InvalidOperationException("剪贴板正被其他程序（" + holder + "）占用，请稍后再试。");
                }
                // Whoever has the clipboard open may be waiting for data that this thread renders
                // (PowerPoint's own copy), so let messages sent to this thread through while we wait.
                MsgWaitForMultipleObjects(0, IntPtr.Zero, false, 20, QS_SENDMESSAGE);
                NativeMessage msg;
                PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_NOREMOVE | PM_QS_SENDMESSAGE);
            }
        }

        /// <summary>Name of the program that has the clipboard open, for the log and the error message.</summary>
        private static string Holder()
        {
            try
            {
                IntPtr w = GetOpenClipboardWindow();
                if (w == IntPtr.Zero) return "未知程序";
                uint pid;
                GetWindowThreadProcessId(w, out pid);
                return System.Diagnostics.Process.GetProcessById((int)pid).ProcessName;
            }
            catch { return "未知程序"; }
        }

        private static bool Skip(uint f)
        {
            // handles to GDI objects: the system recreates bitmaps and metafile pictures from DIB and EMF
            if (f == CF_BITMAP || f == CF_METAFILEPICT || f == CF_PALETTE || f == CF_OWNERDISPLAY ||
                f == CF_DSPBITMAP || f == CF_DSPMETAFILEPICT || f == CF_DSPENHMETAFILE) return true;
            if (f >= CF_PRIVATEFIRST && f <= CF_GDIOBJLAST) return true;
            if (f < 0xC000) return false;
            // OLE's own bookkeeping, our marks, and PowerPoint's "internal" formats: those are pointers into
            // PowerPoint's memory that stop working once the clipboard changes (PowerPoint then pastes from GVML)
            string name = Name(f);
            return name == "DataObject" || name == "Ole Private Data" || name == IgnoreFormat || name == ExcludeFormat ||
                   name.StartsWith("PowerPoint 12.0 Internal", StringComparison.Ordinal);
        }

        private static string Name(uint f)
        {
            var sb = new StringBuilder(256);
            return GetClipboardFormatName(f, sb, sb.Capacity) > 0 ? sb.ToString() : "";
        }

        private static byte[] Read(uint f, ref long budget)
        {
            IntPtr h = GetClipboardData(f);
            if (h == IntPtr.Zero) return null;
            if (f == CF_ENHMETAFILE)
            {
                uint n = GetEnhMetaFileBits(h, 0, null);
                if (n == 0 || (budget -= n) < 0) return null;
                var bits = new byte[n];
                return GetEnhMetaFileBits(h, n, bits) == n ? bits : null;
            }
            long size = (long)GlobalSize(h).ToUInt64();
            if (size <= 0 || (budget -= size) < 0) return null;
            IntPtr p = GlobalLock(h);
            if (p == IntPtr.Zero) return null;
            try
            {
                var data = new byte[size];
                Marshal.Copy(p, data, 0, (int)size);
                return data;
            }
            finally { GlobalUnlock(h); }
        }

        private static void Set(uint f, byte[] data)
        {
            if (f == CF_ENHMETAFILE)
            {
                IntPtr emf = SetEnhMetaFileBits((uint)data.Length, data);
                if (emf != IntPtr.Zero && SetClipboardData(f, emf) == IntPtr.Zero) DeleteEnhMetaFile(emf);
                return;
            }
            IntPtr h = GlobalAlloc(GMEM_MOVEABLE, new UIntPtr((uint)Math.Max(1, data.Length)));
            if (h == IntPtr.Zero) return;
            IntPtr p = GlobalLock(h);
            if (p == IntPtr.Zero) { GlobalFree(h); return; }
            Marshal.Copy(data, 0, p, data.Length);
            GlobalUnlock(h);
            if (SetClipboardData(f, h) == IntPtr.Zero) GlobalFree(h);   // on success the system owns the memory
        }

        private static void MarkPrivate()
        {
            Set(RegisterClipboardFormat(IgnoreFormat), new byte[4]);     // Ditto and other clipboard managers
            Set(RegisterClipboardFormat(ExcludeFormat), new byte[4]);    // Windows clipboard history and cloud clipboard
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeMessage { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public int x, y, pad; }

        [DllImport("user32.dll", SetLastError = true)] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern IntPtr GetOpenClipboardWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseClipboard();
        [DllImport("user32.dll", SetLastError = true)] private static extern bool EmptyClipboard();
        [DllImport("user32.dll", SetLastError = true)] private static extern uint EnumClipboardFormats(uint format);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetClipboardData(uint format);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetClipboardData(uint format, IntPtr mem);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern int GetClipboardFormatName(uint format, StringBuilder name, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern uint RegisterClipboardFormat(string name);
        [DllImport("user32.dll")] private static extern uint MsgWaitForMultipleObjects(uint count, IntPtr handles, bool waitAll, uint ms, uint wakeMask);
        [DllImport("user32.dll")] private static extern bool PeekMessage(out NativeMessage msg, IntPtr hwnd, uint min, uint max, uint remove);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalLock(IntPtr mem);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GlobalUnlock(IntPtr mem);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern UIntPtr GlobalSize(IntPtr mem);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalFree(IntPtr mem);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern uint GetEnhMetaFileBits(IntPtr emf, uint size, byte[] buffer);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr SetEnhMetaFileBits(uint size, byte[] data);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern bool DeleteEnhMetaFile(IntPtr emf);
    }
}
