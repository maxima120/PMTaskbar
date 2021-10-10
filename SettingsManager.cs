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
        public ImageSource imgSrc { get; set; }
        public string link { get; set; }
    }
}
