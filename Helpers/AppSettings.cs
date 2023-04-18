using System;
using System.Configuration;
using System.Linq;

namespace ZapMVImager.Helpers
{
    public static class AppSettings
    {
        public static bool Get(string key, bool value)
        {
            if (!ConfigurationManager.AppSettings.AllKeys.Contains(key) || string.IsNullOrEmpty(ConfigurationManager.AppSettings[key]))
            {
                return value;
            }

            return ConfigurationManager.AppSettings[key].ToLower() == "true";
        }

        public static string Get(string key, string value)
        {
            if (!ConfigurationManager.AppSettings.AllKeys.Contains(key) || string.IsNullOrEmpty(ConfigurationManager.AppSettings[key]))
            {
                return value;
            }

            return ConfigurationManager.AppSettings[key];
        }

        public static int Get(string key, int value)
        {
            if (!ConfigurationManager.AppSettings.AllKeys.Contains(key) || string.IsNullOrEmpty(ConfigurationManager.AppSettings[key]))
            {
                return value;
            }

            return int.Parse(ConfigurationManager.AppSettings[key]);
        }

        public static void AddOrUpdate(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }
    }
}
