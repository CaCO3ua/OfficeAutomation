using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.Factory;

namespace OfficeAutomation
{
    [TestClass]
    public class OpenWordFile
    {
        private Application word;
        private Window mainWindow;
        private Window childWindow;
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
        public void OpenWordDocument()
        {
            clickOpenOtherDocuments();
            clickMore();
            childWindow = mainWindow.ModalWindow("Open");
            insertFilePath();
            clickOpen();
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

        private void insertFilePath()
        {
            childWindow.Get<TextBox>(SearchCriteria.ByAutomationId("1148")).Enter(filePath);
        }

        private void clickOpen()
        {
            IUIItem OkButton = childWindow.Get(SearchCriteria.ByClassName("Button").AndAutomationId("1"));
            OkButton.Click();
        }
    }
}