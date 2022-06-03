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
using System.Windows.Controls.Primitives;
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

        public MainWindow()
        {
            InitializeComponent();

            settingsManager = new SettingsManager<UserSettings>("pmt.settings.json");
            settings = settingsManager.LoadSettings();

            ApplySettings(settings);

            var v = this.GetType().Assembly.GetName().Version.ToString();
            Trace.WriteLine("PMT v." + v);
        }

        private void ApplySettings(UserSettings settings)
        {
            this.Top = settings.Top;
            this.Left = settings.Left;
            this.Height = settings.Height;
            Debug.WriteLine("Init links is starting.");

            var sw = Stopwatch.StartNew();

            var windows = PInvoker.GetTaskBarWindows();

            foreach (var link in settings.Links)
            {
                settings.Items.Add(CreateItem(link, windows));
            }

            Trace.WriteLine($"PMT. Init links processed in: {sw.Elapsed}");

            this.DataContext = settings;
        }

        #endregion

        #region dep prop

        protected override void OnLocationChanged(EventArgs e)
        {
            var currentScreenWidth = SystemParameters.WorkArea.Width;

            // NB: watch the magic numbers :)
            PopupPlacement = currentScreenWidth - this.Left - this.Width - 80 - 20 <= 0 ? PlacementMode.Left : PlacementMode.Right;

            base.OnLocationChanged(e);
        }

        public PlacementMode PopupPlacement
        {
            get { return (PlacementMode)GetValue(PopupPlacementProperty); }
            set { SetValue(PopupPlacementProperty, value); }
        }

        // Using a DependencyProperty as the backing store for PopupPlacement.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty PopupPlacementProperty =
            DependencyProperty.Register("PopupPlacement", typeof(PlacementMode), typeof(MainWindow), new PropertyMetadata(PlacementMode.Right));

        #endregion

        #region lnk hack

        // https://blez.wordpress.com/2013/02/18/get-file-shortcuts-target-with-c/
        // https://github.com/libyal/documentation/blob/main/reference/lnk_the_windows_shortcut_file_format.pdf
        // https://github.com/libyal/liblnk/blob/main/documentation/Windows%20Shortcut%20File%20(LNK)%20format.asciidoc

        // TODO: doesnt work with RDP links to concrete servers
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

        #endregion

        #region UI events

        private void ListView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("System.Windows.Controls.ListViewItem"))
            {
                // this has been processed in the item drop handler
                return;
            }

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

                settings.Items.Add(CreateItem(s));
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
                Process.Start(new ProcessStartInfo { FileName = item.LnkPath, UseShellExecute = true });            
                RefreshItemWindowsAsync(item);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"PMT. Process start exception: {ex}");
                SystemSounds.Exclamation.Play();
            }

            lst.SelectedItem = null;
        }

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.Height <= 2 * 60)
                return;

            this.Height -= 60;

            settings.Height = this.Height;
            settingsManager.SaveSettings(settings);
        }

        private void DnButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.Height >= 20 * 60)
                return;

            this.Height += 60;

            settings.Height = this.Height;
            settingsManager.SaveSettings(settings);
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
                settings.Items.Remove(item);
                settingsManager.SaveSettings(settings);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"PMT. Unpin exception: {ex}");
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
                LnkPath = filename,
                IconImg = PInvoker.GetIcon(filename),
                LnkTarget = link,
                Name = Path.GetFileNameWithoutExtension(link),
            };

            item.Windows = new ObservableCollection<LinkWindow>(GetItemWindows(item, windows));

            return item;
        }

        private List<LinkWindow> GetItemWindows(LinkItem item, List<IntPtr> windows = null)
        {
            var processes = Process.GetProcessesByName(item.Name);

            var result = new List<LinkWindow>();

            if (processes == null || processes.Length == 0)
                return result;

            // TODO: can you make it faster? (init links runs for 140ms and individual links for 5-50ms)

            if (windows == null)
                windows = PInvoker.GetTaskBarWindows();

            var wp = windows.ToDictionary(k => k, v => (int)PInvoker.GetWindowThreadProcessId(v));

            foreach (var process in processes)
            {
                try
                {
                    // TODO: this presumably filters out similar processes started from another source (not LNK)
                    //       but not sure I need it
                    //if (item.LnkTarget == process.MainModule?.FileName)
                    {
                        foreach (var w in wp.Where(i => i.Value == process.Id))
                        {
                            result.Add(new LinkWindow(item) { Process = process, Window = w.Key });
                        }
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

                foreach (var w in item.Windows.ToList())
                {
                    if (!itemWindows.Contains(w))
                        item.Windows.Remove(w);
                }
                foreach (var w in itemWindows.ToList())
                {
                    if (!item.Windows.Contains(w))
                    {
                        w.Title = PInvoker.GetWindowTitle(w.Window);
                        w.ImgSrc = PInvoker.GetWindowThumb(w.Window, 80, 80);
                        item.Windows.Add(w);
                    }
                }

                if (showPopup && item.Windows.Count != 0)
                    item.IsPopupShow = true;
            }

            sw.Stop();
            Debug.WriteLine($"Refresh Item: {sw.Elapsed}");
        }

        private void WindowRestoreButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkWindow;

            if (item == null || item.Window == IntPtr.Zero)
                return;

            PInvoker.SetForegroundWindow(item.Window);
        }

        private void WindowCloseButton_Click(object sender, RoutedEventArgs e)
        {
            var item = (sender as Button).DataContext as LinkWindow;

            if (item == null || item.Window == IntPtr.Zero)
                return;

            //item.process.CloseMainWindow();
            //var b = WindowEnumerator.DestroyWindow(item.window);
            var b = PInvoker.PostMessage(item.Window, (uint)PInvoker.WM.CLOSE, 0, 0);

            RefreshItemWindowsAsync(item.Parent);
        }

        #endregion

        #region item rearrange

        private void ListView_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var point = e.GetPosition(this);
            var item = sender as ListViewItem;
            if (item == null)
                return;

            DragDrop.DoDragDrop(lst, item, DragDropEffects.Move);
        }
        private void ListItem_Drop(object sender, DragEventArgs e)
        {
            try
            {
                var dst = sender as ListViewItem;
                var dstData = dst?.DataContext as LinkItem;

                var dropData = e.Data?.GetData("System.Windows.Controls.ListViewItem");
                var src = dropData as ListViewItem;
                var srcData = src?.DataContext as LinkItem;

                if (dstData != null && srcData != null && srcData != dstData)
                {
                    var dstIdx = settings.Items.IndexOf(dstData);
                    var srcIdx = settings.Items.IndexOf(srcData);

                    if(dstIdx != -1 && srcIdx != -1)
                    {
                        settings.Items.RemoveAt(srcIdx);
                        settings.Items.Insert(dstIdx, srcData);
                        settingsManager.SaveSettings(settings);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        #endregion
    }
}
