using System;
using System.Windows.Automation;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Actions;
using TestStack.White.UIItems.Custom;
using TestStack.White.UIItems.Finders;

namespace OfficeAutomation
{
    [ControlTypeMapping(CustomUIItemType.SplitButton)]
    public class OpenButton : CustomUIItem
    {
        // Implement these two constructors. The order of parameters should be same.
        public OpenButton(AutomationElement automationElement, ActionListener actionListener)
            : base(automationElement, actionListener)
        {
        }

        //Empty constructor is mandatory with protected or public access modifier.
        protected OpenButton() { }

        //Sample method
        public virtual void clickButton()
        {
            //Base class, i.e. CustomUIItem has property called Container. Use this find the items within this.
            //Can also use SearchCriteria for find items
            Container.Get<Button>(SearchCriteria.ByClassName("Button").AndAutomationId("1")).Click();
        }
    }
}
