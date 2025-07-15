using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows;
using NLog;
using NLog.Config;

namespace VOR
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint GetDriveType(string lpRootPathName);

        // Метод для проверки, является ли путь сетевым
        private bool IsNetworkPath(string path)
        {
            string rootPath = Path.GetPathRoot(path);
            uint driveType = GetDriveType(rootPath);

            return (driveType == 4) || rootPath.StartsWith("\\\\"); // DRIVE_REMOTE указывает, что путь находится на сетевом диске
        }
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            string exePath = AppDomain.CurrentDomain.BaseDirectory;

            if (IsNetworkPath(exePath))
            {
                MessageBox.Show("Программа не может быть запущена с сетевого диска", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Environment.Exit(0); // Завершаем приложение
            }

            /*Initialize();*/
        }

        public void Initialize()
        {
            string pathDll = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Console.WriteLine(pathDll);
            var configuration = new LoggingConfiguration();
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = Path.Combine(pathDll, "Logs", $"{DateTime.Now:ddMMyyyy}.log") };
            configuration.AddRule(LogLevel.Trace, LogLevel.Fatal, logfile);

            /*var logconsole = new NLog.Targets.ConsoleTarget("logconsole");
            configuration.AddRule(LogLevel.Trace, LogLevel.Fatal, logconsole);*/

            LogManager.Configuration = configuration;
        }
    }
}
