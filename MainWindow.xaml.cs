using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PMTaskbar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region ctor and fields

        private readonly SettingsManager<UserSettings> settingsManager;
        private readonly UserSettings settings;

        Timer timer;

        public MainWindow()
        {
            InitializeComponent();

            settingsManager = new SettingsManager<UserSettings>("pmt.settings.json");
            settings = settingsManager.LoadSettings();

            ProcessSettings(settings);

            StartWallClock();

            var v = this.GetType().Assembly.GetName().Version.ToString();
            Trace.WriteLine("PMT v." + v);
        }

        #endregion

        #region settings

        private void ProcessSettings(UserSettings settings)
        {
            this.Top = settings.Top;
            this.Left = settings.Left;
            this.Height = settings.Height;

            foreach (var link in settings.Links)
            {
                settings.items.Add(new LinkItem { link = link, imgSrc = GetIcon(link) });
            }

            lst.ItemsSource = settings.items;
        }

        #endregion

        #region pinvokes

        public ImageSource GetIcon(string fileName)
        {
            try
            {
                Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(fileName);

                var imgSrc =

                    System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                            icon.Handle,
                            new Int32Rect(0, 0, icon.Width, icon.Height),
                            BitmapSizeOptions.FromEmptyOptions());

                return imgSrc;

            }
            catch (Exception)
            {
                return null;
            }
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, uint dwNewLong);

        private const int GWL_STYLE = -16;

        private const uint WS_SYSMENU = 0x80000;

        #endregion

        #region overrides

        protected override void OnClosing(CancelEventArgs e)
        {
            settings.Top = this.Top;
            settings.Left = this.Left;
            settings.Height = this.Height;
            settingsManager.SaveSettings(settings);

            base.OnClosing(e);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE,
                GetWindowLong(hwnd, GWL_STYLE) & (0xFFFFFFFF ^ WS_SYSMENU));

            base.OnSourceInitialized(e);
        }

        #endregion

        #region timer

        private void StartWallClock()
        {
            this.TimeText.Text = DateTime.Now.ToString("HH:mm");
            this.WeekdayText.Text = DateTime.Now.ToString("ddd");

            timer = new Timer(1000);
            timer.Start();
            timer.Elapsed += (o, e) =>
            {
                this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    this.TimeText.Text = DateTime.Now.ToString("HH:mm");
                    this.WeekdayText.Text = DateTime.Now.ToString("ddd");
                })
                );
            };

            //Text="{Binding Time, ElementName=clock, StringFormat=\{0:hh\\:mm\}, Mode=OneWay}"
        }

        #endregion

        #region UI events

        private void ListView_Drop(object sender, DragEventArgs e)
        {
            if (!e.Effects.HasFlag(DragDropEffects.Link))
                return;

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            var data = e.Data.GetData(DataFormats.FileDrop);

            if (data == null)
                return;

            var arr = data as string[];

            if (arr == null || arr.Length == 0)
                return;

            for (int i = 0; i < arr.Length; i++)
            {
                var s = arr[i] ?? "";

                if (!s.EndsWith(".lnk"))
                    continue;

                settings.Links.Add(s);
                settings.items.Add(new LinkItem { imgSrc = GetIcon(s), link = s });
                settingsManager.SaveSettings(settings);
            }
        }

        private void lst_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (e.OriginalSource as FrameworkElement)?.DataContext as LinkItem;

            if (item == null)
                return;

            try
            {
                // TODO : see if more needs to be done to avoid any dependency between this process and the children.
                Process.Start(new ProcessStartInfo { FileName = item.link, UseShellExecute = true });
            }
            catch (Exception)
            {
                // TODO
            }

            lst.SelectedItem = null;
        }

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.Height <= 2 * 60)
                return;

            this.Height -= 60;
        }

        private void DnButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.Height >= 20 * 60)
                return;

            this.Height += 60;
        }

        #endregion
    }
}
