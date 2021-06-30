using NUnit.Framework;
using Sessions;

// This is intentionally outside of a namespace as that means NUnit will only run these methods once and not once per namespace.
[SetUpFixture]
// ReSharper disable once CheckNamespace
public class AssemblySetupAndTearDown
{
    [OneTimeSetUp]
    public static void OneTimeSetup()
    {
        // Start the WinAppDriver Session (which may just attach to an existing session)
        WinAppDriverProcess.Instance.StartSession();

        // Start the Desktop Session
        DesktopSession.Instance.StartSession();
    }

    [OneTimeTearDown]
    public static void OneTimeTearDown()
    {
        // Stop the Desktop Session
        DesktopSession.Instance.StopSession();

        // Stop the WinAppDriver (which may kill the process if we started it above)
        WinAppDriverProcess.Instance.StopSession();
    }
}