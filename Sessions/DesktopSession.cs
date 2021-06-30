using System;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;

namespace Sessions
{
    public class DesktopSession : BaseSession<DesktopSession>, IManageSession
    {
        private const string DESKTOP_ID = "Root";

        private static readonly object SessionLock = new();

        private DesktopSession() : base(null, DESKTOP_ID)
        {
        }

        public void StartSession()
        {
            lock (SessionLock)
            {
                if (Session == null)
                {
                    // Create a new session to for the Desktop
                    var opt = new AppiumOptions();
                    opt.AddAdditionalCapability("app", DESKTOP_ID);
                    opt.AddAdditionalCapability("deviceName", "WindowsPC");
                    Session = new WindowsDriver<WindowsElement>(new Uri(WINDOWS_APPLICATION_DRIVER_URL), opt);
                }
            }
        }

        public void StopSession()
        {
            lock (SessionLock)
            {
                Session = null;
            }
        }
    }
}