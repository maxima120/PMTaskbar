using Shell32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Media;

namespace PMTaskbar
{
    public class SettingsManager<T> where T : class, new()
    {
        private readonly string filePath;

        public SettingsManager(string fileName)
        {
            filePath = GetLocalFilePath(fileName);
            Trace.WriteLine(filePath);
        }

        private string GetLocalFilePath(string fileName)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, fileName);
            }
            catch (Exception)
            {
                return fileName;
            }
        }

        public T LoadSettings()
        {
            try
            {
                if (File.Exists(filePath))
                    return JsonSerializer.Deserialize<T>(File.ReadAllText(filePath));
            }
            catch (Exception)
            {
                // TODO
            }

            return new T();
        }

        public void SaveSettings(T settings)
        {
            try
            {
                string json = JsonSerializer.Serialize(settings);
                File.WriteAllText(filePath, json);

            }
            catch (Exception)
            {
                // TODO
            }
        }
    }

    /// <summary>
    /// items holder for serialization
    /// </summary>
    public class UserSettings
    {
        public UserSettings()
        {
            items = new ObservableCollection<LinkItem>();
            Links = new List<string>();
            Top = 10;
            Left = 10;
            Height = 470;
        }

        [JsonIgnore]
        public ObservableCollection<LinkItem> items { get; set; }

        public List<string> Links { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public double Height { get; set; }
    }

    /// <summary>
    /// UI data item
    /// </summary>
    public class LinkItem
    {
        /// <summary>
        /// .lnk icon
        /// </summary>
        public ImageSource imgSrc { get; set; }
        /// <summary>
        /// path to .lnk object (eg ./dektop/my.lnk
        /// </summary>
        public string lnkPath { get; set; }
        /// <summary>
        /// COM lnk object
        /// </summary>
        public ShellLinkObject lnk { get; set; }
        /// <summary>
        /// Target of the link (eg /mypath/myapp.exe)
        /// </summary>
        public string lnkTarget { get; set; }

        /// <summary>
        /// running processes of that target
        /// </summary>
        public List<LinkProcess> processes { get; set; }
    }

    public class LinkProcess
    {
        public string name { get; set; }
        public int processId { get; set; }
    }
}
