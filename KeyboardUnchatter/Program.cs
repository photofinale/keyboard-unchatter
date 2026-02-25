using System;
using System.Windows.Forms;
using System.Windows;
using System.Configuration;
using System.Collections.Specialized;
using System.Diagnostics;
using System.ComponentModel;
using System.Security.Principal;
using System.Threading;
using Microsoft.Win32;

namespace KeyboardUnchatter
{
    static class Program
    {
        private static InputHook _inputHook;
        private static KeyboardMonitor _keyboardMonitor;
        private static MainWindow _mainWindow;
        private static Mutex _singleInstanceMutex;

        #region Get/Set
        public static InputHook InputHook
        {
            get => _inputHook;
        }

        public static KeyboardMonitor KeyboardMonitor
        {
            get => _keyboardMonitor;
        }

        #endregion

        [STAThread]
        static void Main()
        {
            RelaunchAsAdminIfNeeded();

            if (!EnsureSingleInstance())
            {
                return;
            }

            Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.AboveNormal;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (_inputHook = new InputHook())
            {
                // PowerModeChanged abonnieren
                SystemEvents.PowerModeChanged += (s, e) =>
                {
                    if (e.Mode == PowerModes.Resume)
                    {
                        _inputHook.Rehook();
                    }
                };

                _keyboardMonitor = new KeyboardMonitor();

                _mainWindow = new MainWindow();
                _mainWindow.FormClosed += Close;

                if(!Properties.Settings.Default.openMinimized)
                {
                    _mainWindow.Show();
                }

                Application.Run();
            }

            _singleInstanceMutex?.ReleaseMutex();
            _singleInstanceMutex?.Dispose();
        }

        private static bool EnsureSingleInstance()
        {
            bool createdNew;
            _singleInstanceMutex = new Mutex(true, @"Local\KeyboardUnchatter", out createdNew);
            return createdNew;
        }

        public static bool IsRunningAsAdmin()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                if (identity == null)
                {
                    return false;
                }

                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private static void RelaunchAsAdminIfNeeded()
        {
            if (!Properties.Settings.Default.startAsAdmin)
            {
                return;
            }

            RelaunchAsAdmin();
        }

        public static void RelaunchAsAdmin()
        {
            if (IsRunningAsAdmin())
            {
                return;
            }

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = Application.ExecutablePath,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(startInfo);
                Environment.Exit(0);
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
            }
        }

        private static void Close(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private static readonly string runKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

        public static void SetStartup(bool enable)
        {
            string appName = Application.ProductName;
            string exePath = Application.ExecutablePath;

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(runKey, true))
            {
                if (enable)
                {
                    key.SetValue(appName, $"\"{exePath}\"");
                }
                else
                {
                    key.DeleteValue(appName, false);
                }
            }
        }

        public static bool GetStartup()
        {
            // Registry key for user-specific startup programs
            string appName = Application.ProductName;
            string exePath = Application.ExecutablePath;

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(runKey, false))
            {
                if (key == null)
                    return false;

                var value = key.GetValue(appName) as string;
                if (string.IsNullOrEmpty(value))
                    return false;

                // Compare the stored path (trim quotes) with the current executable path
                return string.Equals(
                    value.Trim('"'),
                    exePath,
                    StringComparison.InvariantCultureIgnoreCase
                );
            }
        }
    }
}
