using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace Sessions
{
    public abstract class BaseSession<T> where T : class, IManageSession
    {
        protected const string WINDOWS_APPLICATION_DRIVER_URL = "http://127.0.0.1:4723";
        private static readonly Lazy<T> Lazy = new(CreateInstanceOfT, LazyThreadSafetyMode.ExecutionAndPublication);

        /// <summary>
        /// Constructs a WinAppDriver sessions with the given process and applicationID.
        /// </summary>
        /// <param name="processName"></param>
        /// <param name="applicationId"></param>
        protected BaseSession(string processName, string applicationId)
        {
            this.ProcessName = processName;
            this.ApplicationId = applicationId;
        }

        /// <summary>
        /// Creates an instance of T via reflection since T's constructor is expected to be private.
        /// </summary>
        /// <returns></returns>
        private static T CreateInstanceOfT()
        {
            return Activator.CreateInstance(typeof(T), true) as T;
        }

        protected string ApplicationId { get; }

        protected string ProcessName { get; }

        public static T Instance => Lazy.Value;

        public string MainWindowHandle { get; protected set; }

        public WindowsDriver<WindowsElement> Session { get; protected set; }

        public bool StartedApplication { get; set; }

        public bool IsApplicationStarted
        {
            get
            {
                return Process.GetProcessesByName(ProcessName)
                    .Any(p => p.ProcessName.Equals(ProcessName, StringComparison.InvariantCultureIgnoreCase));
            }
        }

        protected static string GetApplicationPath(string applicationPathKey)
        {
            RegistryKey baseKey = null;

            try // Don't convert to a using block as RegistryKey.OpenBaseKey() can throw an exception.
            {
                baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);

                if (Registry.HasKey(baseKey, applicationPathKey))
                {
                    using var pathKey = Registry.ReadKey(baseKey, applicationPathKey);
                    var path = Registry.ReadValue<string>(pathKey, "Path");
                    var keyName = pathKey.Name.Substring(pathKey.Name.LastIndexOf('\\') + 1);

                    return $"{path}\\{keyName}";
                }
                else
                {
                    throw new InvalidOperationException($"Could not find the path for {applicationPathKey} in the registry.");
                }
            }
            finally
            {
                baseKey?.Dispose();
            }
        }

        protected virtual void CaptureMainWindowHandle()
        {
            // Return all window handles associated with this process/application.
            // At this point hopefully you have one to pick from. Otherwise you can
            // simply iterate through them to identify the one you want.
            var allWindowHandles = Session.WindowHandles;

            // Assuming you only have only one window entry in allWindowHandles and it is in fact the correct one,
            // switch the session to that window as follows. You can repeat this logic with any top window with the same
            // process id (any entry of allWindowHandles)
            MainWindowHandle = allWindowHandles[0]; // Store for later.
            Session.SwitchTo().Window(MainWindowHandle);
        }

        protected virtual void SetApplicationImplicitWaits()
        {
            // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
            Session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
        }

        public virtual void AttachToExistingProcessUnsafe()
        {
            StartedApplication = false;

            foreach (var process in Process.GetProcessesByName(ProcessName).Where(p => p.ProcessName.Equals(ProcessName, StringComparison.InvariantCultureIgnoreCase)))
            {
                var outlookTopLevelWindowHandle = process.MainWindowHandle.ToString("x"); // Convert the handle to a HEX string.

                // Create session by attaching to Outlook top level window
                var opt = new AppiumOptions();
                opt.AddAdditionalCapability("appTopLevelWindow", outlookTopLevelWindowHandle);
                Session = new WindowsDriver<WindowsElement>(new Uri(WINDOWS_APPLICATION_DRIVER_URL), opt);
            }
        }

        public virtual void DetachFromExistingProcessUnsafe()
        {
            // Don't call Session?.Dispose() here as that will quit the app and we only want to detach from it.
            Session = null;
        }

        public void WaitForWindow()
        {
            Thread.Sleep(TimeSpan.FromSeconds(3));
        }

        public IWebDriver SwitchBackToMainWindow()
        {
            WaitForWindow();
            return Session.SwitchTo().Window(MainWindowHandle);
        }

        public IWebDriver SwitchToNewWindow()
        {
            WaitForWindow();
            return Session.SwitchTo().Window(Session.WindowHandles.First());
        }
    }
}
