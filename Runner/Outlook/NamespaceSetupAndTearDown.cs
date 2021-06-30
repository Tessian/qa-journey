using NUnit.Framework;
using Sessions;

namespace Scenarios.Outlook
{
    [SetUpFixture]
    public class NamespaceSetupAndTearDown
    {
        [OneTimeSetUp]
        public static void OneTimeSetup()
        {
            // Create session to launch or bring up application
            OutlookSession.Instance.StartSession();
        }

        [OneTimeTearDown]
        public static void OneTimeTearDown()
        {
            // Close Outlook
            OutlookSession.Instance.StopSession();
        }
    }
}
