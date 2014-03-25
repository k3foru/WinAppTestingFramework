//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class TitlebarActions
    {
        /// <summary>
        /// Click Close method is to click Close on the title bar of all windows
        /// </summary>
        /// <param name="window"></param>
        public static void ClickClose(WinWindow window)
        {
            #region Variable Declarations

            Window.Name = window.Name;
            var uICloseButton = UIEllisWindow.UIEllisTitleBar.UICloseButton;
            #endregion

            // Click 'Close' button
            Mouse.Click(uICloseButton);
        }

        #region Properties
        public static UIEllisWindow UIEllisWindow
        {
            get
            {
                if ((mUIEllisWindow == null))
                {
                    mUIEllisWindow = new UIEllisWindow();
                }
                return mUIEllisWindow;
            }
        }
        #endregion

        #region Fields
        private static UIEllisWindow mUIEllisWindow;
        #endregion
    }

    public class UIEllisWindow : WinWindow
    {

        public UIEllisWindow()
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = Window.Name;
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            WindowTitles.Add(Window.Name);
            #endregion
        }

        #region Properties
        public UIEllisTitleBar UIEllisTitleBar
        {
            get
            {
                if ((mUIEllisTitleBar == null))
                {
                    mUIEllisTitleBar = new UIEllisTitleBar(this);
                }
                return mUIEllisTitleBar;
            }
        }
        #endregion

        #region Fields
        private UIEllisTitleBar mUIEllisTitleBar;
        #endregion
    }


    public class UIEllisTitleBar : WinTitleBar
    {

        public UIEllisTitleBar(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            WindowTitles.Add(Window.Name);
            #endregion
        }

        #region Properties
        public WinButton UICloseButton
        {
            get
            {
                if ((mUICloseButton == null))
                {
                    mUICloseButton = new WinButton(this);
                    #region Search Criteria
                    mUICloseButton.SearchProperties[WinButton.PropertyNames.Name] = "Close";
                    mUICloseButton.WindowTitles.Add(Window.Name);
                    #endregion
                }
                return mUICloseButton;
            }
        }
        #endregion

        #region Fields
        private WinButton mUICloseButton;
        #endregion
    }

    class Window
    {
        public static string Name { get; set; }
    }
}
