using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.Factory;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowStripControls;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems.TabItems;

namespace OfficeAutomation
{
    [TestClass]
    public class OpenOutlookFilePO
    {
        private Application outlook;
        private Window mainWindow;
        string filePath = "C:\\MainDirectory\\Backup.pst";

        [TestInitialize]
        public void Setup()
        {
            outlook = Application.AttachOrLaunch(new System.Diagnostics.ProcessStartInfo("C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE"));
            mainWindow = outlook.GetWindow("Inbox - caco3rv@gmail.com - Outlook", InitializeOption.NoCache);
        }

        [TestCleanup]
        public void TearDown()
        {
            outlook.Close();
        }

        [TestMethod]
        public void OpenOutlookBackupPO()
        {
            clickFile();
            clickOpenExport();
            clickOpenOutlookFile();
            var openFileDialogWindowWrapper = new OpenFileDialogWrapper(mainWindow.ModalWindow("Open Outlook Data File"));
            openFileDialogWindowWrapper.FilePaths.EditableText = filePath;
            openFileDialogWindowWrapper.OkButton.Click();
            Assert.IsNotNull(mainWindow);
        }
        
        private void clickFile()
        {
            mainWindow.Get<Button>(SearchCriteria.ByAutomationId("FileTabButton")).Click();
        }

        private void clickOpenExport()
        {
            mainWindow.Get<TabPage>("Open & Export").Click();
        }

        private void clickOpenOutlookFile()
        {
            mainWindow.Get<Button>("Open Outlook Data File...").Click();
        }
    }
}
