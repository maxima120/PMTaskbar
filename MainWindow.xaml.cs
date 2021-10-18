using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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
            //    XamlWriter.Save(abc.Template, writer);
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

            //lst.ItemsSource = settings.items;
            lst.DataContext = settings;

            WatchProcesses();
        }

        #endregion

        #region pinvokes and lnk hack

        LinkItem CreateItem(string filename)
        {
            // NB: needs to be run as administrator to get the properties
            // Apparently drag drop stops working if ran as administrator.. pfff

            //var sh = new Shell32.Shell();
            //var folder = sh.NameSpace(Path.GetDirectoryName(filename));
            //var folderItem = folder.Items().Item(Path.GetFileName(filename));
            //var link = (Shell32.ShellLinkObject)folderItem.GetLink;

            var link = GetShortcutTarget(filename);

            var o = new LinkItem
            {
                lnkPath = filename,
                imgSrc = GetIcon(filename),
                lnkTarget = link,
                processes = new ObservableCollection<LinkProcess>()
            };

            RefreshItemProcessesAsync(o);

            return o;
        }

        // https://blez.wordpress.com/2013/02/18/get-file-shortcuts-target-with-c/
        // https://github.com/libyal/documentation/blob/main/reference/lnk_the_windows_shortcut_file_format.pdf
        // https://github.com/libyal/liblnk/blob/main/documentation/Windows%20Shortcut%20File%20(LNK)%20format.asciidoc

        // TODO: works for user-created links but not for the "system" RDP link
        private string GetShortcutTarget(string filename)
        {
            try
            {
                if (System.IO.Path.GetExtension(filename).ToLower() != ".lnk")
                    throw new Exception("Supplied file must be a .LNK file");

                var fileStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                using (var fileReader = new BinaryReader(fileStream))
                {
                    var hdr0 = fileReader.ReadUInt32(); // L
                    var hdrGuid = string.Join("", fileReader.ReadBytes(16).Select(i => i.ToString("x2")));
                    var hdrFlags = fileReader.ReadUInt32();
                    var hdrFileAttr = fileReader.ReadUInt32();
                    var hdrTime1 = fileReader.ReadUInt64();
                    var hdrTime2 = fileReader.ReadUInt64();
                    var hdrTime3 = fileReader.ReadUInt64();
                    var hdrLen = fileReader.ReadUInt32(); // lnk file length
                    var hdrIconNum = fileReader.ReadInt32();
                    var hdrShowWnd = fileReader.ReadUInt32();
                    var hdrHotKey = fileReader.ReadUInt16();
                    var hdrUnkn = fileReader.ReadBytes(10);

                    Debug.Assert(fileStream.Position == 0x4c);

                    if ((hdrFlags & 0x01) == 1)
                    {                      
                        // bit 0 set means we have shell item ID list
                        uint offset = fileReader.ReadUInt16();   // Read the length of the Shell item ID list
                        fileStream.Seek(offset, SeekOrigin.Current); // Seek past it (to the file locator info)
                    }

                    long fileInfoStartsAt = fileStream.Position;

                    uint locLen = fileReader.ReadUInt32();   // length of the whole File Location struct
                    uint locAfterOffset = fileReader.ReadUInt32(); // first offset after File Location struct (0x1c)
                    uint locFlags = fileReader.ReadUInt32();
                    uint locVolumeOffset = fileReader.ReadUInt32();
                    uint locBasePathOffset = fileReader.ReadUInt32();
                    uint locNetVolumeOffset = fileReader.ReadUInt32();
                    uint locRemainingPathOffset = fileReader.ReadUInt32();

                    Debug.Assert(fileStream.Position - fileInfoStartsAt == 0x1c);

                    fileStream.Seek((fileInfoStartsAt + locBasePathOffset), SeekOrigin.Begin); // base path
                    long pathLength = (locLen + fileInfoStartsAt) - fileStream.Position - 2; // ignoring 2 x 0x00 terminating chars
                    char[] linkTarget = fileReader.ReadChars((int)pathLength); // should be unicode safe
                    var link = new string(linkTarget);

                    Debug.WriteLine("{0} : {1}", filename, link);
                    //fileStream.Seek(0xc, SeekOrigin.Current); // seek to offset to base pathname
                    //uint fileOffset = fileReader.ReadUInt32(); // read offset to base pathname
                    //                                           // the offset is from the beginning of the file info struct (fileInfoStartsAt)
                    //fileStream.Seek((fileInfoStartsAt + fileOffset), SeekOrigin.Begin); // Seek to beginning of
                    //                                                                    // base pathname (target)
                    //long pathLength = (totalStructLength + fileInfoStartsAt) - fileStream.Position - 2; // read
                    //                                                                                    // the base pathname. I don't need the 2 terminating nulls.
                    //char[] linkTarget = fileReader.ReadChars((int)pathLength); // should be unicode safe
                    //var link = new string(linkTarget);

                    // NB: this doesnt make sense as begin is always 0 but no 00 in the string
                    //int begin = link.IndexOf("\0\0");
                    //if (begin > 0)
                    //    Debugger.Break();

                    //{
                    //    int end = link.IndexOf("\\\\", begin + 2) + 2;
                    //    end = link.IndexOf('\0', end) + 1;

                    //    string firstPart = link.Substring(0, begin);
                    //    string secondPart = link.Substring(end);

                    //    return firstPart + secondPart;
                    //}
                    //else
                    //{
                        return link;
                    //}

                }
            }
            catch(Exception ex)
            {
                return "";
            }
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

            timer = new Timer((o) =>
            {
                this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    this.TimeText.Text = DateTime.Now.ToString("HH:mm");
                    this.WeekdayText.Text = DateTime.Now.ToString("ddd");
                })
                );
            });

            timer.Change(1000 - DateTime.Now.Millisecond, 1000);

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
                RefreshItemProcessesAsync(item);
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

            var errors = 0;

            var sw = Stopwatch.StartNew();

            foreach (var item in settings.items)
            {
                item.processes = new ObservableCollection<LinkProcess>(GetItemProcesses(item));
                //Debug.WriteLine("  process {0} for link {1} is running.", process.Id, item.lnkPath);
            }

            Trace.WriteLine($"Processed in: {sw.Elapsed}, errors: {errors}");
        }

        private List<LinkProcess> GetItemProcesses(LinkItem item)
        {
            var name = Path.GetFileNameWithoutExtension(item.lnkTarget);
            var processes = Process.GetProcessesByName(name);

            if (processes == null || processes.Length == 0)
                return new List<LinkProcess>();
            
            var result = new List<LinkProcess>();

            foreach (var process in processes)
            {
                try
                {
                    if (item.lnkTarget == process.MainModule?.FileName)
                    {
                        result.Add(new LinkProcess(item) { process = process });
                    }
                }
                catch (Win32Exception ex)
                {
                    Debug.WriteLine("Win32Exception {0}:{1}", process.Id, process.ProcessName);
                    continue;
                }
                catch (InvalidOperationException)
                {
                    Debug.WriteLine("InvalidOperationException {0}:{1}", process.Id, process.ProcessName);
                    continue;
                }
                catch (Exception)
                {
                    Debug.WriteLine("Exception {0}:{1}", process.Id, process.ProcessName);
                    continue;
                }
            }

            return result;
        }

        #endregion

        #region mini-windows

        LinkItem itemMouseOver = null;
        private void ListViewItem_MouseEnter(object sender, MouseEventArgs e)
        {
            itemMouseOver = (sender as ListViewItem).DataContext as LinkItem;

            if (itemMouseOver == null)
                return;

            RefreshItemProcessesAsync(itemMouseOver, true);
        }

        private void ListViewItem_MouseLeave(object sender, MouseEventArgs e)
        {
            itemMouseOver = null;

            var item = (sender as ListViewItem).DataContext as LinkItem;

            if (item == null)
                return;

            ThreadPool.QueueUserWorkItem((o) =>
            {
                var item = o as LinkItem;
                Thread.Sleep(400);

                this.Dispatcher.BeginInvoke((Action<LinkItem>)((item) => item.IsPopupShow = false), item);
            },
            item);
        }


        private void RefreshItemProcessesAsync(LinkItem item, bool showPopup = false)
        {
            ThreadPool.QueueUserWorkItem((o) =>
            {
                var data = o as Tuple<LinkItem, bool>;
                Thread.Sleep(400);

                this.Dispatcher.BeginInvoke((Action<LinkItem>)((item) => RefreshItemProcesses(data.Item1, data.Item2)), item);
            },
            new Tuple<LinkItem, bool> (item,showPopup));
        }

        private void RefreshItemProcesses(LinkItem item, bool showPopup)
        {
            var sw = Stopwatch.StartNew();

            lock (item)
            {
                var processes = new ObservableCollection<LinkProcess>(GetItemProcesses(item));

                foreach (var p in item.processes.ToList())
                {
                    if (!processes.Contains(p))
                        item.processes.Remove(p);
                }
                foreach (var p in processes.ToList())
                {
                    if (!item.processes.Contains(p))
                        item.processes.Add(p);
                }

                if (showPopup && item.processes.Count != 0)
                    item.IsPopupShow = true;
            }

            sw.Stop();
            Debug.WriteLine("Refresh Item: {0}", sw.Elapsed);
        }

        public void BringWindowToFront(Process process)
        {
            try
            {
                IntPtr handle = process.MainWindowHandle;

                Debug.WriteLine("Main handle: pid {0} window {1}", process.Id, handle);

                //SetForegroundWindow(handle);
                //ShowWindowAsync(handle, (int)SHOWWINDOW.RESTORE);

                //return;

                // How to find which window is real:
                // https://stackoverflow.com/questions/7277366/why-does-enumwindows-return-more-windows-than-i-expected

                //var children = WindowEnumerator.GetChildWindows(handle);
                //var windows = WindowEnumerator.GetDesktopWindows();
                var windows = WindowEnumerator.GetWindows();

                for (int i = 0; i < windows.Count; i++)
                {
                    var w = (IntPtr)windows[i];

                    uint pid;
                    WindowEnumerator.GetWindowThreadProcessId(w, out pid);

                    if (process.Id == pid)
                    {
                        var v = WindowEnumerator.IsWindowVisible(w);

                        Debug.WriteLine("bingo: pid {0} window {1}, visible {2}", pid, w, v);

                        if (v)
                        {
                            WindowEnumerator.SetForegroundWindow(w);
                            //Thread.Sleep(500);
                            //ShowWindowAsync(w, (int)SHOWWINDOW.RESTORE);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkProcess;

            if (item == null || item.process == null)
                return;

            BringWindowToFront(item.process);
        }

        private void CloseProcessButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkProcess;

            if (item == null || item.process == null)
                return;

            item.process.CloseMainWindow();

            RefreshItemProcessesAsync(item.parent);

        }

        #endregion
    }
}

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

