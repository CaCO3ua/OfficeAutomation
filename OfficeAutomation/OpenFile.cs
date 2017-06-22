using System;
using TestStack.White.Factory;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.ListBoxItems;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WindowStripControls;

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

        public Button CancelButton { get; private set; }
        public IUIItem OkButton { get; private set; }
        public ComboBox FilePaths { get; private set; }
        public ComboBox FileTypeFilter { get; private set; }
        public ToolStrip AddressBar { get; private set; }

        #region Overwrite all the TestStack.White.UIItems.WindowItems.Window virtual functions
        
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
            OkButton = _window.Get(SearchCriteria.ByClassName("Button").AndAutomationId("1"));
            CancelButton = _window.Get<Button>(SearchCriteria.ByClassName("Button").AndAutomationId("2"));
            FilePaths = _window.Get<ComboBox>(SearchCriteria.ByAutomationId("1148"));
            FileTypeFilter = _window.Get<ComboBox>(SearchCriteria.ByAutomationId("1136"));
            AddressBar = _window.Get<ToolStrip>(SearchCriteria.ByClassName("ToolbarWindow32").AndAutomationId("1001"));
        }
    }
}
