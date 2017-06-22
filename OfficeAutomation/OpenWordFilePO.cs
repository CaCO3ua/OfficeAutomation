using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.Factory;

namespace OfficeAutomation
{
    [TestClass]
    public class OpenWordFilePO
    {
        private Application word;
        private Window mainWindow;
        string filePath = "C:\\MainDirectory\\CV_Umanets.docx";

        [TestInitialize]
        public void Setup()
        {
            word = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("winword"));
            mainWindow = word.GetWindow("Word", InitializeOption.NoCache);
        }

        [TestCleanup]
        public void TearDown()
        {
            word.Close();
        }

        [TestMethod]
        public void OpenWordDocumentPO()
        {
            clickOpenOtherDocuments();
            clickMore();
            var openWordFileDialogWindowWrapper = new OpenFileDialogWrapper(mainWindow.ModalWindow("Open"));
            openWordFileDialogWindowWrapper.FilePaths.EditableText = filePath;
            openWordFileDialogWindowWrapper.OkButton.Click();
            Assert.IsNotNull(mainWindow);
        }

        private void clickOpenOtherDocuments()
        {
            mainWindow.Get<Hyperlink>("Open Other Documents").Click();
        }

        private void clickMore()
        {
            mainWindow.Get<Button>("Browse").Click();
        }
    }
}
