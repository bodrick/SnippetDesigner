using System;
using Microsoft.Win32;

namespace RegistryTools
{
    /// <summary>
    /// get the correct data based on the correct vs reg path
    /// </summary>
    public static class RegistryLocations
    {
        public static string GetVisualStudioUserDataPath(string version)
        {
            var location = string.Empty;
            var vsKey = GetVSRegKey(Registry.CurrentUser, version);
            if (vsKey != null)
            {
                location = (string)vsKey.GetValue("VisualStudioLocation", string.Empty);
            }
            return location;
        }

        public static string GetVSInstallDir(string version = "10.0")
        {
            var location = string.Empty;
            var vsKey = GetVSRegKey(Registry.LocalMachine, version);
            if (vsKey != null)
            {
                location = (string)vsKey.GetValue("InstallDir", string.Empty);
            }
            return location;
        }

        public static RegistryKey GetVSRegKey(RegistryKey regKey, string version) => GetVSRegKey(regKey, false, version);

        public static RegistryKey GetVSRegKey(RegistryKey regKey, bool configSection, string version)
        {
            var versionPath = configSection ? version + "_Config" : version;
            var vsKey = regKey.OpenSubKey(@"Software\Microsoft\VisualStudio\" + versionPath);
            if (vsKey == null)
            {
                vsKey = regKey.OpenSubKey(@"Software\Microsoft\VBExpress\" + versionPath);
            }
            if (vsKey == null)
            {
                vsKey = regKey.OpenSubKey(@"Software\Microsoft\VCSExpress\" + versionPath);
            }
            if (vsKey == null)
            {
                vsKey = regKey.OpenSubKey(@"Software\Microsoft\VJSExpress\" + versionPath);
            }
            if (vsKey == null)
            {
                vsKey = regKey.OpenSubKey(@"Software\Microsoft\VCExpress\" + versionPath);
            }
            if (vsKey == null)
            {
                vsKey = regKey.OpenSubKey(@"Software\Microsoft\VWDExpress\" + versionPath);
            }

            return vsKey;
        }

        public static int GetVSUILanguage(string version = "10.0")
        {
            var language = 1033; // default to english
            using (var vsKey = GetVSRegKey(Registry.CurrentUser, false, version))
            {
                if (vsKey != null)
                {
                    using (var generalKey = vsKey.OpenSubKey("General"))
                    {
                        if (generalKey != null)
                        {
                            try
                            {
                                language = (int)vsKey.GetValue("UILanguage", 1033);
                            }
                            catch (InvalidCastException)
                            {
                            }
                        }
                    }
                }
            }

            return language;
        }
    }
}
