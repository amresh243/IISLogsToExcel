// Author: Amresh Kumar (July 2025)

using System.IO;
using System.Windows;

namespace IISLogsToExcel
{
    public partial class App : Application
    {
        private static Mutex? _mutex;

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

            if (e.Args.Length > 0)
            {
                string folderPath = e.Args[0];
                if (Directory.Exists(e.Args[0]))
                {
                    var cmdWindow = new IISLogExporter(folderPath);
                    cmdWindow.Show();
                    return;
                }
            }

            var mainWindow = new IISLogExporter();
            mainWindow.Show();
        }
    }
}
