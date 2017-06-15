using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.WindowsAPI;
using TestStack.White.InputDevices;
using TestStack.White.UIItems.WindowItems;
using System.Collections.Generic;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.ListBoxItems;
using System.Windows.Automation;

namespace OfficeAutomation
{
    [TestClass]
    public class UnitTest2
    {
        private Application word;
        private Window mainWindow;
        private List<Window> modalWindows;
        private Window childWindow;
        string filePath = "C:\\MainDirectory\\CV_Umanets.docx";

        [TestInitialize]
        public void Setup()
        {
            word = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("winword"));
            mainWindow = word.GetWindows().First();
            modalWindows = mainWindow.ModalWindows();
        }

        [TestCleanup]
        public void TearDown()
        {
            word.Close();
        }

        [TestMethod]
        public void OpenWord()
        {
            clickOpenOtherDocuments();
            clickMore();
            childWindow = mainWindow.ModalWindow("Открытие документа");
            insertFilePath();
            clickOpen();
            System.Console.WriteLine(mainWindow.TitleBar);
            Assert.AreEqual(mainWindow.TitleBar, "CV_Umanets");
        }

        private void clickOpenOtherDocuments()
        {
            mainWindow.Get<Hyperlink>("Открыть другие документы").Click();
        }

        private void clickMore()
        {
            mainWindow.Get<Button>("Обзор").Click();
        }

        private void insertFilePath()
        {
            childWindow.Get<TextBox>("Имя файла:").Enter(filePath);
        }

        private void clickOpen()
        {
            Button openButton = (Button)childWindow.Get(SearchCriteria.ByControlType(ControlType.SplitButton).AndAutomationId("1"));
            openButton.Click();
            Menu menuItem = childWindow.Get<Menu>(SearchCriteria.ByText("Открыть"));
            menuItem.Click();
        }
    }
}