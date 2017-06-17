using System;
using TestStack.White.Factory;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems.WindowItems;

namespace OfficeAutomation
{
    internal class OpenFileDialogWrapper : Window
    {
        private readonly Window _window;

        public OpenFileDialogWrapper(Window window)
        {
            _window = window;
            LoadControls();
        }

        public TestStack.White.UIItems.Button CancelButton { get; private set; }
        public IUIItem OkButton { get; private set; }
        public TestStack.White.UIItems.ListBoxItems.ComboBox FilePaths { get; private set; }
        public TestStack.White.UIItems.ListBoxItems.ComboBox FileTypeFilter { get; private set; }
        public TestStack.White.UIItems.WindowStripControls.ToolStrip AddressBar { get; private set; }

        #region Overwrite all the TestStack.White.UIItems.WindowItems.Window virtual functions

        // For Example
        public override string Title { get { return _window.Title; } }

        public override PopUpMenu Popup => throw new NotImplementedException();

        public override Window ModalWindow(string title, InitializeOption option)
        {
            throw new NotImplementedException();
        }

        public override Window ModalWindow(SearchCriteria searchCriteria, InitializeOption option)
        {
            throw new NotImplementedException();
        }

        #endregion

        private void LoadControls()
        {
            OkButton = _window.Get(SearchCriteria.ByAutomationId("1"));
            CancelButton = _window.Get<TestStack.White.UIItems.Button>(SearchCriteria.ByClassName("Button").AndAutomationId("2"));
            FilePaths = _window.Get<TestStack.White.UIItems.ListBoxItems.ComboBox>(SearchCriteria.ByAutomationId("1148"));
            FileTypeFilter = _window.Get<TestStack.White.UIItems.ListBoxItems.ComboBox>(SearchCriteria.ByAutomationId("1136"));
            AddressBar = _window.Get<TestStack.White.UIItems.WindowStripControls.ToolStrip>(SearchCriteria.ByClassName("ToolbarWindow32").AndAutomationId("1001"));
        }
    }
}
