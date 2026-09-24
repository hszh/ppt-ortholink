using System;
using System.IO;

namespace OrthoLink
{
    /// <summary>Tiny file logger. There is no debugger attached inside PowerPoint, so this is our eyes.</summary>
    internal static class Log
    {
        private static readonly object Gate = new object();
        public static readonly string Path = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OrthoLink", "ortholink.log");

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path));
                    File.AppendAllText(Path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff ") + message + Environment.NewLine);
                }
            }
            catch { /* logging must never throw */ }
        }

        public static void Error(string where, Exception ex)
        {
            Write(where + " FAILED: " + ex.GetType().Name + ": " + ex.Message + Environment.NewLine + ex.StackTrace);
        }
    }
}
