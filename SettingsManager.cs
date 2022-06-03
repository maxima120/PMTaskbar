using Shell32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PMTaskbar
{
    public class SettingsManager<T> where T : IUserSettings, new()
    {
        private readonly string filePath;

        public SettingsManager(string fileName)
        {
            filePath = GetLocalFilePath(fileName);
            Trace.WriteLine($"PMT. Settings: {filePath}");
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
                {
                    return JsonSerializer.Deserialize<T>(File.ReadAllText(filePath));
                }
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
                settings.Unload();
                string json = JsonSerializer.Serialize(settings);
                File.WriteAllText(filePath, json);
            }
            catch (Exception)
            {
                // TODO
            }
        }
    }

    public interface IUserSettings
    {
        void Unload();
    }
    /// <summary>
    /// items holder for serialization
    /// </summary>
    public class UserSettings : IUserSettings
    {
        public UserSettings()
        {
            Items = new ObservableCollection<LinkItem>();
            Links = new List<string>();
            Top = 10;
            Left = 10;
            Height = 470;
        }

        [JsonIgnore]
        public ObservableCollection<LinkItem> Items { get; set; }

        public List<string> Links { get; set; }
        public double Top { get; set; }
        public double Left { get; set; }
        public double Height { get; set; }
        public bool PanelShowSeconds { get; set; }
        public bool PanelShowDate { get; set; }

        public void Unload()
        {
            Links = Items.Select(i => i.LnkPath).ToList();
        }
    }

    /// <summary>
    /// UI data item
    /// </summary>
    public class LinkItem : INotifyPropertyChanged
    {
        public LinkItem()
        {
            isPopupShow = false;
        }

        public string Name { get; set; }

        /// <summary>
        /// .lnk icon
        /// </summary>
        public ImageSource IconImg { get; set; }
        /// <summary>
        /// path to .lnk object (eg ./dektop/my.lnk
        /// </summary>
        public string LnkPath { get; set; }
        ///// <summary>
        ///// COM lnk object
        ///// </summary>
        //public ShellLinkObject lnk { get; set; }
        /// <summary>
        /// Target of the link (eg /mypath/myapp.exe)
        /// </summary>
        public string LnkTarget { get; set; }

        private bool isPopupShow;
        public bool IsPopupShow { get => isPopupShow; set { isPopupShow = value; OnPropertyChanged("IsPopupShow"); } }

        /// <summary>
        /// running processes of that target
        /// </summary>
        public ObservableCollection<LinkWindow> Windows { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(name));
        }
    }

    public class LinkWindow : INotifyPropertyChanged
    {
        public LinkWindow(LinkItem parent)
        {
            Parent = parent;
        }
        
        public IntPtr Window { get; set; }
        public Process Process { get; set; }
        public LinkItem Parent { get; set; }

        private BitmapSource imgSrc;
        public BitmapSource ImgSrc { get => imgSrc; set { if (imgSrc == value) return; imgSrc = value; OnPropertyChanged("ImgSrc"); } }

        private string title;
        public string Title { get => title; set { title = value; OnPropertyChanged("Title"); } }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(name));
        }
    }
}
