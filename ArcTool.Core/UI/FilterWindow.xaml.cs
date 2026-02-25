using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Input;

namespace ArcTool.UI
{
    public class FilterItem : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        public bool IsSelected { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsVisible { get; set; }
        public object Data { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class ViewItem : INotifyPropertyChanged
    {
        private string _viewName = string.Empty;
        public string ViewName
        {
            get => _viewName;
            set { _viewName = value; OnPropertyChanged(); }
        }

        public int FilterCount { get; set; }
        public bool IsSelected { get; set; }
        public object Data { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class FilterWindow : Window, INotifyPropertyChanged
    {
        private string _activeViewName = "Đang tải...";
        public string ActiveViewName
        {
            get => _activeViewName;
            set { _activeViewName = value; OnPropertyChanged(); }
        }

        public ObservableCollection<FilterItem> FiltersSource { get; set; } = new ObservableCollection<FilterItem>();
        public ObservableCollection<ViewItem> ViewsSource { get; set; } = new ObservableCollection<ViewItem>();

        public FilterWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            dgFiltersCopy.ItemsSource = FiltersSource;
            dgViewsPaste.ItemsSource = ViewsSource;

            // Cho phép kéo thả cửa sổ dù đang tương tác với Revit
            this.MouseDown += (s, e) => { if (e.LeftButton == MouseButtonState.Pressed) DragMove(); };
        }

        // Cập nhật tên View an toàn từ luồng Revit sang luồng UI
        public void UpdateActiveViewInfo(string viewName)
        {
            this.Dispatcher.Invoke(() => {
                ActiveViewName = viewName;
                if (txtActiveViewName != null) txtActiveViewName.Text = viewName;
            });
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // Điều khiển cửa sổ
        private void BtnMinimize_Click(object sender, RoutedEventArgs e) => this.WindowState = WindowState.Minimized;
        private void BtnMaximize_Click(object sender, RoutedEventArgs e) => this.WindowState = (this.WindowState == WindowState.Maximized) ? WindowState.Normal : WindowState.Maximized;
        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Hide(); // Ẩn thay vì đóng để giữ sự kiện

        private void BtnAddFilters_Click(object sender, RoutedEventArgs e) { this.DialogResult = true; this.Close(); }
        private void BtnDeleteFilters_Click(object sender, RoutedEventArgs e) { }
        private void BtnCopyFilters_Click(object sender, RoutedEventArgs e) { this.DialogResult = false; this.Close(); }

        private void BtnViewTemplates_Click(object sender, RoutedEventArgs e)
        {
            btnViewTemplates.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E5F1FF"));
            btnViewTemplates.FontWeight = FontWeights.Bold;
            btnViewsSheets.Background = System.Windows.Media.Brushes.White;
            btnViewsSheets.FontWeight = FontWeights.Normal;
        }

        private void BtnViewsSheets_Click(object sender, RoutedEventArgs e)
        {
            btnViewsSheets.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E5F1FF"));
            btnViewsSheets.FontWeight = FontWeights.Bold;
            btnViewTemplates.Background = System.Windows.Media.Brushes.White;
            btnViewTemplates.FontWeight = FontWeights.Normal;
        }
    }
}