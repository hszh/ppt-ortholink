using System;
using Microsoft.Win32;

namespace OrthoLink
{
    /// <summary>User defaults, persisted under HKCU\Software\OrthoLink.</summary>
    internal static class Settings
    {
        private const string KeyPath = @"Software\OrthoLink";

        public static readonly float[] RadiusOptions = { 0f, 4f, 6f, 8f, 12f, 16f, 24f };

        public static float RadiusPt = 12f;
        public static float GapPt = 0f;
        public static bool ArrowEnd = true;
        public static bool ArrowStart = false;
        public static bool AutoFollow = true;

        public static void Load()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return;
                    RadiusPt = ReadFloat(k, "RadiusPt", RadiusPt);
                    GapPt = ReadFloat(k, "GapPt", GapPt);
                    ArrowEnd = ReadBool(k, "ArrowEnd", ArrowEnd);
                    ArrowStart = ReadBool(k, "ArrowStart", ArrowStart);
                    AutoFollow = ReadBool(k, "AutoFollow", AutoFollow);
                }
            }
            catch (Exception ex) { Log.Error("Settings.Load", ex); }
        }

        public static void Save()
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    k.SetValue("RadiusPt", RadiusPt.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    k.SetValue("GapPt", GapPt.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    k.SetValue("ArrowEnd", ArrowEnd ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("ArrowStart", ArrowStart ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("AutoFollow", AutoFollow ? 1 : 0, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex) { Log.Error("Settings.Save", ex); }
        }

        public static int RadiusIndex() { return NearestRadiusIndex(RadiusPt); }

        public static int NearestRadiusIndex(float radiusPt)
        {
            int best = 0;
            for (int i = 1; i < RadiusOptions.Length; i++)
                if (Math.Abs(RadiusOptions[i] - radiusPt) < Math.Abs(RadiusOptions[best] - radiusPt)) best = i;
            return best;
        }

        private static float ReadFloat(RegistryKey k, string name, float fallback)
        {
            var v = k.GetValue(name) as string;
            float f;
            return v != null && float.TryParse(v, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out f) ? f : fallback;
        }

        private static bool ReadBool(RegistryKey k, string name, bool fallback)
        {
            var v = k.GetValue(name);
            return v is int ? ((int)v) != 0 : fallback;
        }
    }
}
