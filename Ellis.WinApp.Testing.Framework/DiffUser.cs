//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.Drawing;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace Ellis.WinApp.Testing.Framework
{
    public class DiffUser
    {
        public static void EllisDifferentUser()
        {
            #region Variable Declarations

            WinListItem uIEllisListItem = UIProgramManagerWindow.UIDesktopList.UIEllisListItem;
            WinMenuItem uIRunasdifferentuserMenuItem = UIItemWindow.UIContextMenu.UIRunasdifferentuserMenuItem;
            WinEdit uIUsernameEdit = UIWindowsSecurityWindow.UIUseanotheraccountListItem.UIUsernameEdit;
            WinEdit uIPasswordEdit = UIWindowsSecurityWindow.UIUseanotheraccountListItem.UIPasswordEdit;
            WinButton uIOKButton = UIWindowsSecurityWindow.UIWindowsSecurityPane.UIOKButton;

            #endregion

            var clickPoints = new Point(uIEllisListItem.BoundingRectangle.Width / 2 + uIEllisListItem.BoundingRectangle.X,
               uIEllisListItem.BoundingRectangle.Height / 2 + uIEllisListItem.BoundingRectangle.Y);

            Mouse.Click(uIEllisListItem, MouseButtons.Right, ModifierKeys.Shift, new Point(clickPoints.X, clickPoints.Y));

            clickPoints = new Point(uIRunasdifferentuserMenuItem.BoundingRectangle.Width / 2 + uIRunasdifferentuserMenuItem.BoundingRectangle.X,
               uIRunasdifferentuserMenuItem.BoundingRectangle.Height / 2 + uIRunasdifferentuserMenuItem.BoundingRectangle.Y);

            Mouse.Click(uIRunasdifferentuserMenuItem, new Point(clickPoints.X, clickPoints.Y));
            Playback.PlaybackSettings.ContinueOnError = true;

            uIUsernameEdit.Text = EllisDifferentUserParams.UIUsernameEditText;
            Keyboard.SendKeys(uIUsernameEdit, EllisDifferentUserParams.UIUsernameEditSendKeys, ModifierKeys.None);
            Keyboard.SendKeys(uIPasswordEdit, EllisDifferentUserParams.UIPasswordEditSendKeys, true);

            clickPoints = new Point(uIOKButton.BoundingRectangle.Width / 2 + uIOKButton.BoundingRectangle.X,
               uIOKButton.BoundingRectangle.Height / 2 + uIOKButton.BoundingRectangle.Y);

            Mouse.Click(uIOKButton, new Point(clickPoints.X, clickPoints.Y));
            Playback.PlaybackSettings.ContinueOnError = false;
        }

        public static void EllisDifferentUser(string username, string password)
        {
            #region Variable Declarations

            WinListItem uIEllisListItem = UIProgramManagerWindow.UIDesktopList.UIEllisListItem;
            WinMenuItem uIRunasdifferentuserMenuItem = UIItemWindow.UIContextMenu.UIRunasdifferentuserMenuItem;
            WinEdit uIUsernameEdit = UIWindowsSecurityWindow.UIUseanotheraccountListItem.UIUsernameEdit;
            WinEdit uIPasswordEdit = UIWindowsSecurityWindow.UIUseanotheraccountListItem.UIPasswordEdit;
            WinButton uIOKButton = UIWindowsSecurityWindow.UIWindowsSecurityPane.UIOKButton;

            #endregion

            var clickPoints = new Point(uIEllisListItem.BoundingRectangle.Width / 2 + uIEllisListItem.BoundingRectangle.X,
               uIEllisListItem.BoundingRectangle.Height / 2 + uIEllisListItem.BoundingRectangle.Y);

            Mouse.Click(uIEllisListItem, MouseButtons.Right, ModifierKeys.Shift, new Point(clickPoints.X, clickPoints.Y));

            clickPoints = new Point(uIRunasdifferentuserMenuItem.BoundingRectangle.Width / 2 + uIRunasdifferentuserMenuItem.BoundingRectangle.X,
               uIRunasdifferentuserMenuItem.BoundingRectangle.Height / 2 + uIRunasdifferentuserMenuItem.BoundingRectangle.Y);

            Mouse.Click(uIRunasdifferentuserMenuItem, new Point(clickPoints.X, clickPoints.Y));
            Playback.PlaybackSettings.ContinueOnError = true;

            uIUsernameEdit.Text = username;
            Keyboard.SendKeys(uIUsernameEdit, EllisDifferentUserParams.UIUsernameEditSendKeys, ModifierKeys.None);
            Keyboard.SendKeys(uIPasswordEdit, password, false);

            clickPoints = new Point(uIOKButton.BoundingRectangle.Width / 2 + uIOKButton.BoundingRectangle.X,
               uIOKButton.BoundingRectangle.Height / 2 + uIOKButton.BoundingRectangle.Y);

            Mouse.Click(uIOKButton, new Point(clickPoints.X, clickPoints.Y));
            Playback.PlaybackSettings.ContinueOnError = false;
        }

        #region Properties

        public static EllisDifferentUserParams EllisDifferentUserParams
        {
            get
            {
                if ((mEllisDifferentUserParams == null))
                {
                    mEllisDifferentUserParams = new EllisDifferentUserParams();
                }
                return mEllisDifferentUserParams;
            }
        }

        public static UIProgramManagerWindow UIProgramManagerWindow
        {
            get
            {
                if ((mUIProgramManagerWindow == null))
                {
                    mUIProgramManagerWindow = new UIProgramManagerWindow();
                }
                return mUIProgramManagerWindow;
            }
        }

        public static UIItemWindow UIItemWindow
        {
            get
            {
                if ((mUIItemWindow == null))
                {
                    mUIItemWindow = new UIItemWindow();
                }
                return mUIItemWindow;
            }
        }

        public static UIWindowsSecurityWindow UIWindowsSecurityWindow
        {
            get
            {
                if ((mUIWindowsSecurityWindow == null))
                {
                    mUIWindowsSecurityWindow = new UIWindowsSecurityWindow();
                }
                return mUIWindowsSecurityWindow;
            }
        }

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

        public static UIItemWindow1 UIItemWindow1
        {
            get
            {
                if ((mUIItemWindow1 == null))
                {
                    mUIItemWindow1 = new UIItemWindow1();
                }
                return mUIItemWindow1;
            }
        }

        #endregion

        #region Fields

        private static EllisDifferentUserParams mEllisDifferentUserParams;
        private static UIProgramManagerWindow mUIProgramManagerWindow;
        private static UIItemWindow mUIItemWindow;
        private static UIWindowsSecurityWindow mUIWindowsSecurityWindow;
        private static UIEllisWindow mUIEllisWindow;
        private static UIItemWindow1 mUIItemWindow1;

        #endregion
    }

    public class EllisDifferentUserParams
    {

        #region Fields

        public string UIUsernameEditText = "EllisCSR";
        public string UIUsernameEditSendKeys = "{Tab}";
        public string UIPasswordEditSendKeys = "TjjlAvmdqfc1APOXfIe+x399zNhSUqeX";

        #endregion
    }

    public class UIProgramManagerWindow : WinWindow
    {
        public UIProgramManagerWindow()
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Program Manager";
            SearchProperties[PropertyNames.ClassName] = "Progman";
            WindowTitles.Add("Program Manager");
            #endregion
        }

        #region Properties
        public UIDesktopList UIDesktopList
        {
            get
            {
                if ((mUIDesktopList == null))
                {
                    mUIDesktopList = new UIDesktopList(this);
                }
                return mUIDesktopList;
            }
        }

        private UIDesktopList mUIDesktopList;
        #endregion
    }

    public class UIDesktopList : WinList
    {

        public UIDesktopList(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Desktop";
            WindowTitles.Add("Program Manager");
            #endregion
        }

        #region Properties
        public WinListItem UIEllisListItem
        {
            get
            {
                if ((mUIEllisListItem == null))
                {
                    mUIEllisListItem = new WinListItem(this);
                    #region Search Criteria
                    mUIEllisListItem.SearchProperties[WinListItem.PropertyNames.Name] = "Ellis";
                    mUIEllisListItem.WindowTitles.Add("Program Manager");
                    #endregion
                }
                return mUIEllisListItem;
            }
        }

        private WinListItem mUIEllisListItem;
        #endregion
    }

    public class UIItemWindow : WinWindow
    {
        public UIItemWindow()
        {
            #region Search Criteria
            SearchProperties[PropertyNames.AccessibleName] = "Context";
            SearchProperties[PropertyNames.ClassName] = "#32768";
            #endregion
        }

        #region Properties
        public UIContextMenu UIContextMenu
        {
            get
            {
                if ((mUIContextMenu == null))
                {
                    mUIContextMenu = new UIContextMenu(this);
                }
                return mUIContextMenu;
            }
        }

        private UIContextMenu mUIContextMenu;
        #endregion
    }

    public class UIContextMenu : WinMenu
    {
        public UIContextMenu(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Context";
            #endregion
        }

        #region Properties
        public WinMenuItem UIRunasdifferentuserMenuItem
        {
            get
            {
                if ((mUIRunasdifferentuserMenuItem == null))
                {
                    mUIRunasdifferentuserMenuItem = new WinMenuItem(this);
                    #region Search Criteria
                    mUIRunasdifferentuserMenuItem.SearchProperties[WinMenuItem.PropertyNames.Name] = "Run as different user";
                    #endregion
                }
                return mUIRunasdifferentuserMenuItem;
            }
        }

        private WinMenuItem mUIRunasdifferentuserMenuItem;
        #endregion
    }

    public class UIWindowsSecurityWindow : WinWindow
    {
        public UIWindowsSecurityWindow()
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Windows Security";
            SearchProperties[PropertyNames.ClassName] = "#32770";
            WindowTitles.Add("Windows Security");
            #endregion
        }

        #region Properties
        public UIUseanotheraccountListItem UIUseanotheraccountListItem
        {
            get
            {
                if ((mUIUseanotheraccountListItem == null))
                {
                    mUIUseanotheraccountListItem = new UIUseanotheraccountListItem(this);
                }
                return mUIUseanotheraccountListItem;
            }
        }

        public UIWindowsSecurityPane UIWindowsSecurityPane
        {
            get
            {
                if ((mUIWindowsSecurityPane == null))
                {
                    mUIWindowsSecurityPane = new UIWindowsSecurityPane(this);
                }
                return mUIWindowsSecurityPane;
            }
        }

        private UIUseanotheraccountListItem mUIUseanotheraccountListItem;
        private UIWindowsSecurityPane mUIWindowsSecurityPane;
        #endregion
    }

    public class UIUseanotheraccountListItem : WinListItem
    {
        public UIUseanotheraccountListItem(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Use another account";
            WindowTitles.Add("Windows Security");
            #endregion
        }

        #region Properties
        public WinEdit UIUsernameEdit
        {
            get
            {
                if ((mUIUsernameEdit == null))
                {
                    mUIUsernameEdit = new WinEdit(this);
                    #region Search Criteria
                    mUIUsernameEdit.SearchProperties[WinEdit.PropertyNames.Name] = "User name";
                    mUIUsernameEdit.WindowTitles.Add("Windows Security");
                    #endregion
                }
                return mUIUsernameEdit;
            }
        }

        public WinEdit UIPasswordEdit
        {
            get
            {
                if ((mUIPasswordEdit == null))
                {
                    mUIPasswordEdit = new WinEdit(this);
                    #region Search Criteria
                    mUIPasswordEdit.SearchProperties[WinEdit.PropertyNames.Name] = "Password";
                    mUIPasswordEdit.WindowTitles.Add("Windows Security");
                    #endregion
                }
                return mUIPasswordEdit;
            }
        }

        private WinEdit mUIUsernameEdit;
        private WinEdit mUIPasswordEdit;
        #endregion
    }

    public class UIWindowsSecurityPane : WinPane
    {
        public UIWindowsSecurityPane(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Windows Security";
            WindowTitles.Add("Windows Security");
            #endregion
        }

        #region Properties
        public WinButton UIOKButton
        {
            get
            {
                if ((mUIOKButton == null))
                {
                    mUIOKButton = new WinButton(this);
                    #region Search Criteria
                    mUIOKButton.SearchProperties[WinButton.PropertyNames.Name] = "OK";
                    mUIOKButton.WindowTitles.Add("Windows Security");
                    #endregion
                }
                return mUIOKButton;
            }
        }

        private WinButton mUIOKButton;
        #endregion
    }

    public class UIEllisWindow : WinWindow
    {
        public UIEllisWindow()
        {
            #region Search Criteria
            SearchProperties[PropertyNames.Name] = "Ellis";
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            WindowTitles.Add("Ellis");
            #endregion
        }

        #region Properties
        public UI_ShellLayoutView_TooMenuBar UI_ShellLayoutView_TooMenuBar
        {
            get
            {
                if ((mUI_ShellLayoutView_TooMenuBar == null))
                {
                    mUI_ShellLayoutView_TooMenuBar = new UI_ShellLayoutView_TooMenuBar(this);
                }
                return mUI_ShellLayoutView_TooMenuBar;
            }
        }

        private UI_ShellLayoutView_TooMenuBar mUI_ShellLayoutView_TooMenuBar;
        #endregion
    }

    public class UI_ShellLayoutView_TooMenuBar : WinMenuBar
    {
        public UI_ShellLayoutView_TooMenuBar(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties[WinMenu.PropertyNames.Name] = "Toolbar";
            WindowTitles.Add("Ellis");
            #endregion
        }

        #region Properties
        public WinMenuItem UIFileMenuItem
        {
            get
            {
                if ((mUIFileMenuItem == null))
                {
                    mUIFileMenuItem = new WinMenuItem(this);
                    #region Search Criteria
                    mUIFileMenuItem.SearchProperties[WinMenuItem.PropertyNames.Name] = "File";
                    mUIFileMenuItem.WindowTitles.Add("Ellis");
                    #endregion
                }
                return mUIFileMenuItem;
            }
        }

        private WinMenuItem mUIFileMenuItem;
        #endregion
    }

    public class UIItemWindow1 : WinWindow
    {
        public UIItemWindow1()
        {
            #region Search Criteria
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            SearchProperties[PropertyNames.Instance] = "4";
            #endregion
        }

        #region Properties
        public UIItemWindow2 UIItemWindow
        {
            get
            {
                if ((mUIItemWindow == null))
                {
                    mUIItemWindow = new UIItemWindow2(this);
                }
                return mUIItemWindow;
            }
        }

        private UIItemWindow2 mUIItemWindow;
        #endregion
    }

    public class UIItemWindow2 : WinWindow
    {
        public UIItemWindow2(UITestControl searchLimitContainer) :
            base(searchLimitContainer)
        {
            #region Search Criteria
            SearchProperties.Add(new PropertyExpression(PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            #endregion
        }

        #region Properties
        public WinMenu UIItemMenu
        {
            get
            {
                if ((mUIItemMenu == null))
                {
                    mUIItemMenu = new WinMenu(this);
                }
                return mUIItemMenu;
            }
        }

        private WinMenu mUIItemMenu;
        #endregion
    }
}
