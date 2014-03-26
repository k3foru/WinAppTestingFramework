//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using Microsoft.VisualStudio.TestTools.UITesting;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class MouseActions : AppContext
    {
        public static void Click(UITestControl control)
        {
            Actions.MouseClickOnCoordinates(control);
            Actions.SendEnter();
        }

        public static void DoubleClick(UITestControl control)
        {
            Mouse.DoubleClick(control);
        }

        public static void ClickButton(UITestControl windowProperties, string control)
        {
            Mouse.Click(Actions.GetWindowChild(windowProperties, control));
        }
    }
}
