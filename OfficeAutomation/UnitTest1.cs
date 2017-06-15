using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using System.Linq;

namespace OfficeAutomation
{
    [TestClass]
    public class UnitTest1
    {
        private Application notepad;

        [TestInitialize]
        public void Setup()
        {
            notepad = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("notepad"));
        }

        [TestCleanup]
        public void TearDown()
        {
            notepad.Close();
        }

        [TestMethod]
        public void TestMethod1()
        {
            var mainWindow = notepad.GetWindows().First();
            Assert.IsNotNull(mainWindow);
        }
    }
}