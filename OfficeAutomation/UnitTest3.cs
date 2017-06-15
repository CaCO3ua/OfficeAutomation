using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.WindowsAPI;
using TestStack.White.InputDevices;
using TestStack.White.UIItems.WindowItems;
using System.Collections.Generic;

namespace OfficeAutomation
{
    [TestClass]
    public class UnitTest3
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
        public void OpenWordPO()
        {
            clickOpenOtherDocuments();
            clickMore();
            var openFileDialogWindowWrapper = new OpenFileDialogWrapper(mainWindow.ModalWindow("Открытие документа"));
            openFileDialogWindowWrapper.FilePaths.EditableText = filePath;
            openFileDialogWindowWrapper.OkButton.clickButton();
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
            childWindow.Get<Button>("1").Click();
        }
    }
}
