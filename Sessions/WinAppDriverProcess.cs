using System.Diagnostics;

namespace Sessions
{
    public class WinAppDriverProcess : BaseSession<WinAppDriverProcess>, IManageSession
    {
        private const string WIN_APP_DRIVER_PATH = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        private const string PROCESS_NAME = "WinAppDriver";
        private static Process _winAppDriverProcess;
        private static readonly object WinAppDriverProcessLock = new();

        public WinAppDriverProcess() : base(PROCESS_NAME, WIN_APP_DRIVER_PATH)
        { }

        public void StartSession()
        {
            lock (WinAppDriverProcessLock)
            {
                if (_winAppDriverProcess == null)
                {
                    if (IsApplicationStarted)
                        AttachToExistingProcessUnsafe();
                    else
                        StartApplicationUnsafe();
                }
            }
        }

        public void StopSession()
        {
            lock (WinAppDriverProcessLock)
            {
                if (_winAppDriverProcess == null)
                    return;

                // Only stop the application if we didn't attach to an existing one.
                if (StartedApplication)
                    StopApplicationUnsafe();
                else
                    DetachFromExistingProcessUnsafe();
            }
        }

        public override void AttachToExistingProcessUnsafe()
        {
            StartedApplication = false;
        }

        public void StartApplicationUnsafe()
        {
            _winAppDriverProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = WIN_APP_DRIVER_PATH,
                }
            };

            _winAppDriverProcess.Start();
            StartedApplication = true;
        }

        public override void DetachFromExistingProcessUnsafe()
        {
            // Nothing to do here as we didn't start the process.
        }

        public void StopApplicationUnsafe()
        {
            try
            {
                if (!_winAppDriverProcess.HasExited)
                {
                    _winAppDriverProcess?.Kill();
                }

                _winAppDriverProcess?.Dispose();
            }
            finally
            {
                _winAppDriverProcess = null;
            }
        }
    }
}