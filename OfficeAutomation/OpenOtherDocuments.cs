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
    internal class OpenOtherDocumentsWrapper : Window
    {
        private readonly Window _window;

        public OpenOtherDocumentsWrapper(Window window)
        {
            _window = window;
            LoadControls();
        }

        public Hyperlink OpenOtherDocuments { get; private set; }
        public Button Browse { get; private set; }

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
            OpenOtherDocuments = _window.Get<Hyperlink>(SearchCriteria.ByText("Open Other Documents"));
            Browse = _window.Get<Button>(SearchCriteria.ByText("Browse"));

        }
    }
}
