using System;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace Sessions
{
    public sealed class OutlookSession : BaseSession<OutlookSession>, IManageSession
    {
        private const string APPLICATION_PATH_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";
        private const string PROCESS_NAME = "OUTLOOK";

        private static readonly object SessionLock = new();

        private OutlookSession() : base(PROCESS_NAME, GetApplicationPath(APPLICATION_PATH_KEY))
        { }

        public void StartSession()
        {
            lock (SessionLock)
            {
                if (Session == null)
                {
                    if (IsApplicationStarted)
                    {
                        AttachToExistingProcessUnsafe();
                    }
                    else
                    {
                        StartApplicationUnsafe();
                    }

                    CaptureMainWindowHandle();
                    SetApplicationImplicitWaits();
                }
            }
        }

        public void StopSession()
        {
            lock (SessionLock)
            {
                if (Session == null)
                    return;

                // Wait for things to send.
                Instance.WaitForSend();
                Session.SwitchTo().Window(MainWindowHandle);

                // Only stop the application if we didn't attach to an existing one.
                if (StartedApplication)
                {
                    StopApplicationUnsafe();
                }
                else
                {
                    DetachFromExistingProcessUnsafe();
                }
            }
        }

        public void StartApplicationUnsafe()
        {
            // Launch application if it is not yet launched
            if (Session == null)
            {
                StartedApplication = true;

                // Create a new session to launch Outlook application
                var opt = new AppiumOptions();
                opt.AddAdditionalCapability("app", base.ApplicationId);
                opt.AddAdditionalCapability("deviceName", "WindowsPC");
                Session = new WindowsDriver<WindowsElement>(new Uri(WINDOWS_APPLICATION_DRIVER_URL), opt);

                // Identify the current window handle, this will be the splash screen. You can check through inspect.exe which window this is.
                var currentWindowHandle = Session.CurrentWindowHandle;

                // Wait for however long it is needed for the right window to appear/for the splash screen to be dismissed
                while (Session.WindowHandles.Any(wh => wh.Equals(currentWindowHandle)))
                {
                    Thread.Sleep(TimeSpan.FromMilliseconds(100));
                }
            }
        }

        public void StopApplicationUnsafe()
        {
            try
            {
                var windowHandle = Session.WindowHandles[0];

                do
                {
                    Session.SwitchTo().Window(windowHandle);
                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    Session.Close();

                    // If the main window is prompting about sending an email... let it send.
                    WaitForSendDialog();

                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    windowHandle = Session.WindowHandles[0] ?? null;
                } while (!string.IsNullOrWhiteSpace(windowHandle));
            }
            catch
            {
                // purposely swallowing any exceptions here.
            }
            finally
            {
                Session?.Quit();
                Session?.Dispose();
                Thread.Sleep(TimeSpan.FromSeconds(2));
                Session = null;
            }
        }

        public void WaitForSend()
        {
            Thread.Sleep(TimeSpan.FromSeconds(3));
        }

        public void WaitForSendDialog()
        {
            try
            {
                Thread.Sleep(TimeSpan.FromMilliseconds(2));

                WindowsElement waitForSend = null;

                do
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));

                    waitForSend = Session.FindElement(By.Name("Don't Exit"));
                } while (waitForSend != null);

            }
            catch
            {
                // Swallow any exceptions as if this dialog doesn't occur we get an exception.
            }
        }
    }
}