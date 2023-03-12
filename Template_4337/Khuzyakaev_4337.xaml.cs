using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace Template_4337
{
    public partial class Khuzyakaev_4337 : Window
    {
        public Khuzyakaev_4337()
        {
            InitializeComponent();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            // for .NET Core you need to add UseShellExecute = true
            // see https://learn.microsoft.com/dotnet/api/system.diagnostics.processstartinfo.useshellexecute#property-value
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}