using System;
using System.Collections.Generic;
using System.Linq;
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
