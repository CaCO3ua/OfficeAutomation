using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.Factory;

namespace OfficeAutomation
{
    [TestClass]
    public class OpenExcelFilePO
    {
        private Application excel;
        private Window mainWindow;
        string filePath = "C:\\MainDirectory\\Document.xlsx";

        [TestInitialize]
        public void Setup()
        {
            excel = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("excel"));
            mainWindow = excel.GetWindow("Excel", InitializeOption.NoCache);
        }

        [TestCleanup]
        public void TearDown()
        {
            excel.Close();
        }

        [TestMethod]
        public void OpenExcelWorkbookPO()
        {
            clickOpenOtherWorkbooks();
            clickMore();
            var openFileDialogWindowWrapper = new OpenFileDialogWrapper(mainWindow.ModalWindow("Open"));
            openFileDialogWindowWrapper.FilePaths.EditableText = filePath;
            openFileDialogWindowWrapper.OkButton.Click();
            Assert.IsNotNull(mainWindow);
        }

        private void clickOpenOtherWorkbooks()
        {
            mainWindow.Get<Hyperlink>("Open Other Workbooks").Click();
        }

        private void clickMore()
        {
            mainWindow.Get<Button>("Browse").Click();
        }
    }
}
