using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.Factory;

namespace OfficeAutomation
{
    [TestClass]
    public class OpenPPDFilePO
    {
        private Application powerPoint;
        private Window mainWindow;
        string filePath = "C:\\MainDirectory\\Presentation.pptx";

        [TestInitialize]
        public void Setup()
        {
            powerPoint = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("powerpnt"));
            mainWindow = powerPoint.GetWindow("PowerPoint", InitializeOption.NoCache);
        }

        [TestCleanup]
        public void TearDown()
        {
            powerPoint.Close();
        }

        [TestMethod]
        public void OpenPPPresentationPO()
        {
            clickOpenOtherPresentations();
            clickMore();
            var openFileDialogWindowWrapper = new OpenFileDialogWrapper(mainWindow.ModalWindow("Open"));
            openFileDialogWindowWrapper.FilePaths.EditableText = filePath;
            openFileDialogWindowWrapper.OkButton.Click();
            Assert.IsNotNull(mainWindow);
        }

        private void clickOpenOtherPresentations()
        {
            mainWindow.Get<Hyperlink>("Open Other Presentations").Click();
        }

        private void clickMore()
        {
            mainWindow.Get<Button>("Browse").Click();
        }
    }
}
