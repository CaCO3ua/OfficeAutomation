using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.WindowsAPI;
using TestStack.White.UIItems.WindowItems;
using System.Collections.Generic;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.ListBoxItems;
using System.Windows.Automation;
using TestStack.White.Factory;
using System.Reflection;
using TestStack.White.UIItems.Custom;

namespace OfficeAutomation
{
    [TestClass]
    public class UnitTest2
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
            IUIItem openButton = childWindow.Get(SearchCriteria.ByAutomationId("1"));
            openButton.Click();
        }
    }
}