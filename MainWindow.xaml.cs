using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

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

            // Useful - extract default template without mucking with Blend

            //var str = new StringBuilder();
            //using (var writer = new StringWriter(str))
            //    XamlWriter.Save(btn.Template, writer);
            //Debug.Write(str);
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
                settings.items.Add(CreateItem(link));
            }

            lst.ItemsSource = settings.items;

            WatchProcesses();
        }

        #endregion

        #region pinvokes

        LinkItem CreateItem(string filename)
        {
            // NB: needs to be run as administrator to get the properties

            var sh = new Shell32.Shell();
            var folder = sh.NameSpace(Path.GetDirectoryName(filename));
            var folderItem = folder.Items().Item(Path.GetFileName(filename));
            var link = (Shell32.ShellLinkObject)folderItem.GetLink;

            var o = new LinkItem { lnkPath = filename, imgSrc = GetIcon(filename), lnk = link, processes = new List<LinkProcess>() };

            //[17080] Target: System.__ComObject
            //try
            //{
            //    Debug.WriteLine("Path: " + link.Path);
            //}
            //catch { }
            //try
            //{
            //    Debug.WriteLine("WorkingDirectory: " + link.WorkingDirectory);
            //}
            //catch { }
            //try
            //{
            //    Debug.WriteLine("Arguments: " + link.Arguments);
            //}
            //catch { }
            //try
            //{
            //    Debug.WriteLine("Description: " + link.Description);
            //}
            //catch { }

            // test link
            try
            {
                Debug.WriteLine("Path: " + link.Path);
                o.lnkTarget = link.Path;
            }
            catch 
            {
                // lnk is unusable
                o.lnk = null;
                o.lnkTarget = null;
            }

            return o;
        }

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
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            var data = e.Data.GetData(DataFormats.FileDrop);

            if (data == null)
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            var arr = data as string[];

            if (arr == null || arr.Length == 0)
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            var playSound = false;
            for (int i = 0; i < arr.Length; i++)
            {
                var s = arr[i] ?? "";

                if (!s.EndsWith(".lnk"))
                {
                    playSound = true;
                    continue;
                }
                if (settings.Links.Contains(s))
                {
                    playSound = true;
                    continue;
                }

                settings.Links.Add(s);
                settings.items.Add(CreateItem(s));
                settingsManager.SaveSettings(settings);
            }

            if (playSound)
            {
                SystemSounds.Hand.Play();
            }
        }

        private void lst_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = (e.OriginalSource as FrameworkElement)?.DataContext as LinkItem;

            if (item == null)
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            try
            {
                // TODO : see if more needs to be done to avoid any dependency between this process and the children.
                Process.Start(new ProcessStartInfo { FileName = item.lnkPath, UseShellExecute = true });
            }
            catch (Exception)
            {
                // TODO
                SystemSounds.Exclamation.Play();
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

        private void UnpinMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var item = lst.SelectedItem as LinkItem;

            if (item == null)
            {
                SystemSounds.Exclamation.Play();
                return;
            }

            try
            {
                settings.Links.Remove(item.lnkPath);
                settings.items.Remove(item);
                settingsManager.SaveSettings(settings);

                lst.Items.Remove(item);
            }
            catch (Exception)
            {
                // TODO
                SystemSounds.Exclamation.Play();
            }
        }

        #endregion

        #region process watch

        // https://www.codeproject.com/Articles/12138/Process-Information-and-Notifications-using-WMI
        // https://www.codeproject.com/Tips/44329/Edit-shortcuts-lnk-properties-with-C

        // https://stackoverflow.com/questions/3556048/how-to-detect-win32-process-creation-termination-in-c/50315772#50315772
        // https://social.msdn.microsoft.com/Forums/vstudio/en-US/1c82bfb2-7c90-4b08-b34d-e64d1b9af006/wmi-event-watcher-query-causes-high-cpu-usage-for-wmiprvseexe-and-slowdown?forum=netfxbcl

        void WatchProcesses()
        {
            Debug.WriteLine("WatchProcesses is starting.");

            var sw = Stopwatch.StartNew();

            // NB: WMI timing is x10 of the .NET GetProcesses

            // TODO : try select with WHERE and specific names
            //var queryString = @"SELECT Name, ProcessId, ExecutablePath FROM Win32_Process WHERE ExecutablePath = 'C:\Users\Public\Desktop\TablePlus.lnk'";

            //var searcher = new ManagementObjectSearcher(@"\\.\root\CIMV2", queryString);
            //var processes = searcher.Get();

            //foreach (var process in processes)
            //{
            //    var name = process["Name"].ToString();
            //    var processId = Convert.ToInt32(process["ProcessId"]);
            //    var executablePath = process["ExecutablePath"]?.ToString() ?? "";

            //    var item = settings.items.SingleOrDefault(i => i.lnkTarget == executablePath);

            //    if (item == null)
            //        continue;

            //    Debug.WriteLine("  process {0} for link {1} is running.", processId, item.lnkPath);
            //}

            //Debug.WriteLine("Processed in: {0}", sw.Elapsed);

            //sw.Restart();
            //Process[] pps = Process.GetProcesses();

            // NB: without stopwords - there are exceptions which make it run for 500-600ms instead of 10ms
            //var stopWords = new[] {
            //    "Idle",
            //    "System",
            //    "Registry",
            //    "smss",
            //    "csrss",
            //    "wininit",
            //    "csrss",
            //    "services",
            //    "Memory Compression",
            //    "MBAMService",
            //    "svchost",
            //    "SecurityHealthService",
            //    "SgrmBroker",
            //    "svchost"
            //};

            var errors = 0;
            //foreach (var process in pps)
            //{
            //    //Debug.WriteLine("{0}:{1}", process.Id, process.ProcessName);

            //    //if (stopWords.Contains(process.ProcessName))
            //    //    continue;

            //    string module = null;
            //    try
            //    {
            //        module = process.MainModule?.FileName;
            //    }
            //    catch (Win32Exception ex)
            //    {
            //        errors++;
            //        Debug.WriteLine("Win32Exception {0}:{1}", process.Id, process.ProcessName);
            //        continue;
            //    }
            //    catch (InvalidOperationException)
            //    {
            //        errors++;
            //        Debug.WriteLine("InvalidOperationException {0}:{1}", process.Id, process.ProcessName);
            //        continue;
            //    }
            //    catch (Exception)
            //    {
            //        errors++;
            //        Debug.WriteLine("Exception {0}:{1}", process.Id, process.ProcessName);
            //        continue;
            //    }

            //    // TODO : index by target and PID
            //    var item = settings.items.SingleOrDefault(i => i.lnkTarget == module);

            //    if (item == null)
            //        continue;

            //    Debug.WriteLine("  process {0} for link {1} is running.", process.Id, item.lnkPath);
            //}

            foreach (var item in settings.items)
            {
                var name = Path.GetFileNameWithoutExtension(item.lnkPath);
                var processes = Process.GetProcessesByName(name);

                if (processes != null && processes.Length != 0)
                {
                    foreach (var process in processes)
                    {
                        try
                        {

                            if (item.lnkTarget == process.MainModule?.FileName)
                            {
                                Debug.WriteLine("  process {0} for link {1} is running.", process.Id, item.lnkPath);
                                item.processes.Add(new LinkProcess { name = process.ProcessName, processId = process.Id });
                            }
                        }
                        catch (Win32Exception ex)
                        {
                            errors++;
                            //Debug.WriteLine("Win32Exception {0}:{1}", process.Id, process.ProcessName);
                            continue;
                        }
                        catch (InvalidOperationException)
                        {
                            errors++;
                            Debug.WriteLine("InvalidOperationException {0}:{1}", process.Id, process.ProcessName);
                            continue;
                        }
                        catch (Exception)
                        {
                            errors++;
                            Debug.WriteLine("Exception {0}:{1}", process.Id, process.ProcessName);
                            continue;
                        }
                    }
                }
            }

            Trace.WriteLine($"Processed in: {sw.Elapsed}, errors: {errors}");
        }

        #endregion
    }
}
