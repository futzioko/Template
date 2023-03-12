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

namespace Template_4332
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
        private void Khamitova_4332_Click(object sender, RoutedEventArgs e)
        {
            _4332_Khamitova elina= new _4332_Khamitova();
            elina.Show();
        }

        private void Spiridonov_4332_Click(object sender, RoutedEventArgs e)
        {
            _4332_Spiridonov form = new _4332_Spiridonov();
            form.Show();
        }
    }
}
