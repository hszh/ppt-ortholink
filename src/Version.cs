using System.Reflection;

// Single source of truth for the version. Compiled into both the add-in and the setup program.
[assembly: AssemblyVersion(OrthoLink.Ver.Assembly)]
[assembly: AssemblyFileVersion(OrthoLink.Ver.Assembly)]
[assembly: AssemblyInformationalVersion(OrthoLink.Ver.Display)]

namespace OrthoLink
{
    internal static class Ver
    {
        public const string Display = "1.0.0";
        public const string Assembly = "1.0.0.0";
    }
}
