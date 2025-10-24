using Microsoft.UI.Xaml;
using System.Diagnostics;
using System.Reflection;

namespace Visitor.Card.QR
{
    public partial class App : Application
    {
        private Window? _window;
        public static string AppVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;

        public App()
        {
            InitializeComponent();
        }
        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            _window = new MainWindow();
            _window.Activate();
        }
    }
}
