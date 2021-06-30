using System;
using Microsoft.Win32;

namespace Sessions
{
    public static class Registry
    {
        public static RegistryKey ReadKey(RegistryKey root, string path, RegistryView view = RegistryView.Default, bool isWritable = false)
        {
            RegistryKey registryKey = null;
            try // Don't convert to a using block as root.OpenSubKey() can thrown an exception.
            {
                // Open a subKey as read-only
                registryKey = root.OpenSubKey(path, isWritable);

                return registryKey;
            }
            catch (Exception)
            {
                registryKey?.Dispose();
                return null;
            }
            
        }

        public static T ReadValue<T>(RegistryKey key, string name)
        {
            try
            {
                var value = key.GetValue(name, default);
                return (T) value;
            }
            catch
            {
                return default;
            }
        }

        public static void SetValue<T>(RegistryKey key, string name, T value)
        {
            key.SetValue(name, value, RegistryValueKind.DWord);
        }

        public static bool HasKey(RegistryKey root, string path, RegistryView view = RegistryView.Default)
        {
            using (var ret = root.OpenSubKey(path))
            {
                return ret != null;
            }
        }
    }
}
