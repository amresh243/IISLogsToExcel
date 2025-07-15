// Author: Amresh Kumar (July 2025)

using System.IO;
using System.Windows;

#pragma warning disable CS8618

namespace IISLogsToExcel
{
    public partial class App : Application
    {
        private static Mutex _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            const string mutexName = "IISLogsToExcel";
            bool isNewInstance = false;

            _mutex ??= new Mutex(true, mutexName, out isNewInstance);
            if (!isNewInstance)
            {
                // Another instance is already running
                MessageBox.Show("One instance of application IISLogsToExcel.exe is already running.", "IIS Logs to Excel Converter",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
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
