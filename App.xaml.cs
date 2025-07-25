// Author: Amresh Kumar (July 2025)

using System.IO;
using System.Windows;

namespace IISLogsToExcel
{
    public partial class App : Application
    {
        private static Mutex? _mutex;
        private IISLogExporter? _mainWindow;

        /// Command line support, allows for single instance application only
        protected override void OnStartup(StartupEventArgs e)
        {
            bool isNewInstance = false;

            _mutex ??= new Mutex(true, Constants.ApplicationName, out isNewInstance);
            if (!isNewInstance)
            {
                // Another instance is already running
                MessageBox.Show(Messages.InstanceWarning, Captions.InstanceWarning, MessageBoxButton.OK, MessageBoxImage.Warning);
                Shutdown();
                return;
            }

            base.OnStartup(e);

            _mainWindow = (e.Args.Length > 0 && Directory.Exists(e.Args[0]))
                ? new IISLogExporter(e.Args[0])
                : new IISLogExporter();

            _mainWindow.Show();
        }
    }
}
