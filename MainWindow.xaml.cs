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

        private void ProcessSettings(UserSettings settings)
        {
            this.Top = settings.Top;
            this.Left = settings.Left;
            this.Height = settings.Height;

            Debug.WriteLine("Init links is starting.");

            var sw = Stopwatch.StartNew();

            var windows = PInvoker.GetTaskBarWindows();

            foreach (var link in settings.Links)
            {
                settings.items.Add(CreateItem(link, windows));
            }

            Trace.WriteLine($"Init links processed in: {sw.Elapsed}");

            lst.DataContext = settings;
        }

        #endregion

        #region lnk hack

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
            PInvoker.SetWindowLong(hwnd, PInvoker.GWL_STYLE, 
                PInvoker.GetWindowLong(hwnd, PInvoker.GWL_STYLE) & (0xFFFFFFFF ^ PInvoker.WS_SYSMENU));

            base.OnSourceInitialized(e);
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
                RefreshItemWindowsAsync(item);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Process start exception: {ex}");
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

        LinkItem CreateItem(string filename, List<IntPtr> windows = null)
        {
            var link = GetShortcutTarget(filename);

            var item = new LinkItem
            {
                lnkPath = filename,
                imgSrc = GetIcon(filename),
                lnkTarget = link,
                name = Path.GetFileNameWithoutExtension(link),
            };

            item.windows = new ObservableCollection<LinkWindow>(GetItemWindows(item, windows));

            return item;
        }

        private List<LinkWindow> GetItemWindows(LinkItem item, List<IntPtr> windows = null)
        {
            var processes = Process.GetProcessesByName(item.name);

            var result = new List<LinkWindow>();

            if (processes == null || processes.Length == 0)
                return result;

            // TODO: can you make it faster? (init links runs for 180ms and individual links for 5-45ms)

            if (windows == null)
                windows = PInvoker.GetTaskBarWindows();

            var wp = windows.ToDictionary(k => k, v => (int)PInvoker.GetWindowThreadProcessId(v));

            foreach (var process in processes)
            {
                try
                {
                    // TODO: this presumably filters out similar processes started from another source (not LNK)
                    //       but not sure I need it
                    if (item.lnkTarget == process.MainModule?.FileName)
                    {
                        foreach (var w in wp.Where(i => i.Value == process.Id))
                            result.Add(new LinkWindow(item) { process = process, window = w.Key });
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

            RefreshItemWindowsAsync(itemMouseOver, true);
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

        private void RefreshItemWindowsAsync(LinkItem item, bool showPopup = false, List<IntPtr> windows = null, int delay = 400)
        {
            ThreadPool.QueueUserWorkItem(data =>
            {
                Thread.Sleep(delay);

                this.Dispatcher.BeginInvoke((Action<LinkItem>)((item) => RefreshItemWindows(data.item, data.flag, data.w)), item);
            },
            (item: item, flag: showPopup, w: windows), 
            false);
        }

        private void RefreshItemWindows(LinkItem item, bool showPopup, List<IntPtr> windows)
        {
            var sw = Stopwatch.StartNew();

            lock (item)
            {
                var itemWindows = GetItemWindows(item, windows);

                foreach (var p in item.windows.ToList())
                {
                    if (!itemWindows.Contains(p))
                        item.windows.Remove(p);
                }
                foreach (var p in itemWindows.ToList())
                {
                    if (!item.windows.Contains(p))
                        item.windows.Add(p);
                }

                if (showPopup && item.windows.Count != 0)
                    item.IsPopupShow = true;
            }

            sw.Stop();
            Debug.WriteLine("Refresh Item: {0}", sw.Elapsed);
        }

        private void WindowRestoreButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkWindow;

            if (item == null || item.window == IntPtr.Zero)
                return;

            PInvoker.SetForegroundWindow(item.window);
        }

        private void WindowCloseButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkWindow;

            if (item == null || item.window == IntPtr.Zero)
                return;

            //item.process.CloseMainWindow();
            //var b = WindowEnumerator.DestroyWindow(item.window);
            var b = PInvoker.PostMessage(item.window, (uint)PInvoker.WM.CLOSE, 0, 0);

            RefreshItemWindowsAsync(item.parent);
        }

        #endregion
    }
}

// NB: needs to be run as administrator to get the properties
// Apparently drag drop stops working if ran as administrator.. pfff

//var sh = new Shell32.Shell();
//var folder = sh.NameSpace(Path.GetDirectoryName(filename));
//var folderItem = folder.Items().Item(Path.GetFileName(filename));
//var link = (Shell32.ShellLinkObject)folderItem.GetLink;


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

