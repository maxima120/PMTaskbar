using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Interaction logic for TaskbarPanel.xaml
    /// </summary>
    public partial class TaskbarPanel : UserControl
    {
        public TaskbarPanel()
        {
            InitializeComponent();
            StartWallClock();
        }

        #region depprop

        public bool ShowSeconds
        {
            get { return (bool)GetValue(ShowSecondsProperty); }
            set { SetValue(ShowSecondsProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ShowSeconds.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ShowSecondsProperty =
            DependencyProperty.Register("ShowSeconds", typeof(bool), typeof(TaskbarPanel), new PropertyMetadata(false, OnShowSecondsChanged));

        static void OnShowSecondsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var c = d as TaskbarPanel;
            var val = (bool)e.NewValue;
            c.timeFormat = val ? "HH:mm:ss" : "HH:mm";
        }

        public bool ShowDate
        {
            get { return (bool)GetValue(ShowDateProperty); }
            set { SetValue(ShowDateProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ShowDate.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ShowDateProperty =
            DependencyProperty.Register("ShowDate", typeof(bool), typeof(TaskbarPanel), new PropertyMetadata(false, OnShowDateChanged));


        static void OnShowDateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var c = d as TaskbarPanel;
            var val = (bool)e.NewValue;
            c.DateBorder.Visibility = val ? Visibility.Visible : Visibility.Collapsed;
        }

        #endregion


        #region timer

        Timer timer;

        string timeFormat = "HH:mm";
        private void StartWallClock()
        {
            Tick();

            timer = new Timer((o) =>
            {
                this.Dispatcher.BeginInvoke((Action)(() => Tick()));
                timer.Change(1000 - DateTime.Now.Millisecond + 10, 1000);
            });

            timer.Change(1000 - DateTime.Now.Millisecond + 10, 1000);
        }

        private void Tick()
        {
            this.TimeText.Text = DateTime.Now.ToString(timeFormat);
            this.WeekdayText.Text = DateTime.Now.ToString("ddd");
            if(ShowDate)
                this.DateText.Text = DateTime.Now.ToString("dd MMM");
        }

        #endregion

        private void TimeText_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ShowSeconds = !ShowSeconds;
        }

        private void DateText_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ShowDate = !ShowDate;
        }
    }
}
