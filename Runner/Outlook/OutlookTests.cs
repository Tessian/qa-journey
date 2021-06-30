using System.Linq;
using NUnit.Framework;
using OpenQA.Selenium;
using Sessions;

namespace Scenarios.Outlook
{
    [TestFixture]
    public class OutlookTests
    {
        [SetUp]
        public void BeforeEachTest()
        {
            // Always start each test from the main application window.
            OutlookSession.Instance.SwitchBackToMainWindow();
        }

        [TearDown]
        public void AfterEachTest()
        {
            // Always end each test from the main application window.
            OutlookSession.Instance.SwitchBackToMainWindow();
        }

        [Test]
        public void ComposeNewEmail()
        {
            // Get the "New Email" button
            var newEmailButton = OutlookSession.Instance.Session.FindElement(By.Name("New Email"));

            // Click it
            newEmailButton.Click();

            // switch to the new window
            var composeWindow = OutlookSession.Instance.SwitchToNewWindow();

            // Get the various fields to fill out.
            var toField = composeWindow
                .FindElements(By.Name("To")) // There are two elements named "To", so get the one that is of the correct type.
                .First(e => e.TagName.Equals("ControlType.Edit"));
            var subjectField = composeWindow
                .FindElements(By.Name("Subject"))
                .First(e => e.TagName.Equals("ControlType.Edit"));
            var bodyField = composeWindow // There's only one element named "Untitled Message", so no need to filter by type.
                .FindElement(By.Name("Untitled Message"));

            // Now let's type into the above fields:
            toField.SendKeys("address1@domain.tld");
            subjectField.SendKeys("This is my subject.");
            bodyField.SendKeys("This is the body of the email.  I can type a lot in here!");

            // Let's get the send button and click on it.
            composeWindow.FindElements(By.Name("Send"))
                .First(e => e.TagName.Equals("ControlType.Button"))
                .Click();
        }
    }
}
