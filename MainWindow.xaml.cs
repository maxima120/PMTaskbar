using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
        public MainWindow()
        {
            InitializeComponent();

            items = new ObservableCollection<LinkItem>();
            lst.ItemsSource = items;
        }

        ObservableCollection<LinkItem> items;

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

                Debug.WriteLine("GOT: " + s);

                items.Add(new LinkItem { imgSrc = GetIcon(s), link = s });
            }
        }

        private void ListView_MouseMove(object sender, MouseEventArgs e)
        {
            //Debug.WriteLine("{0}", e.LeftButton);
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

        protected override void OnSourceInitialized(EventArgs e)
        {
            IntPtr hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE,
                GetWindowLong(hwnd, GWL_STYLE) & (0xFFFFFFFF ^ WS_SYSMENU));

            base.OnSourceInitialized(e);
        }

        private void lst_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = ((FrameworkElement)e.OriginalSource).DataContext as LinkItem;
            if (item == null)
                return;

            Debug.WriteLine("CLICK: " + item.link);

            Process.Start(new ProcessStartInfo { FileName = item.link, UseShellExecute = true });;

            lst.SelectedItem = null;
        }
    }

    public class LinkItem
    {
        public ImageSource imgSrc { get; set; }
        public string link { get; set; }
    }
}
