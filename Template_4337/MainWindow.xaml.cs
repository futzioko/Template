using System.Windows;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Khuzyakaev_4337_Click(object sender, RoutedEventArgs e)
        {
            var window = new Khuzyakaev_4337();
            window.Show();
        }

        private void Gumerov_4337_Click(object sender, RoutedEventArgs e)
        {
            var window = new _4337_Gumerov();
            window.Show();
        }
    }
}
