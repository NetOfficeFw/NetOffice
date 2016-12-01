using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Runtime;
using System.Collections;

namespace NetOffice.OfficeApi.Tools.Utils
{       
    /// <summary>
    /// Tray related utils
    /// </summary>
    public class TrayUtils
    {
        #region Fields

        private CommonUtils _owner;
        private NotifyIcon _icon;
        private Icon _applicationIcon;
        private bool _applicationIconResolved;
        private bool _visible = true;
        private bool _suspendOnAutomation = true;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal TrayUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
            Menu = new TrayMenu(owner.Owner);             
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs when the balloon tip is clicked
        /// </summary>
        public event EventHandler BalloonTipClicked;

        /// <summary>
        /// Occurs when the balloon tip is closed by the user
        /// </summary>
        public event EventHandler BalloonTipClosed;
        
        /// <summary>
        /// Occurs when the balloon tip is displayed on the screen
        /// </summary>
        public event EventHandler BalloonTipShown;
        
        /// <summary>
        /// Occurs when the user clicks the icon in the notification area
        /// </summary>
        public event EventHandler Click;
        
        /// <summary>
        /// Occurs when the user double-clicks the icon in the notification area of the taskbar
        /// </summary>
        public event EventHandler DoubleClick;
        
        /// <summary>
        /// Occurs when the user clicks a NotifyIcon with the mouse
        /// </summary>
        public event ToolsMouseEventHandler MouseClick;
        
        /// <summary>
        /// Occurs when the user double-clicks the NotifyIcon with the mouse
        /// </summary>
        public event ToolsMouseEventHandler MouseDoubleClick;
        
        /// <summary>
        /// Occurs when the user presses the mouse button while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event ToolsMouseEventHandler MouseDown;
        
        /// <summary>
        /// Occurs when the user moves the mouse while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event ToolsMouseEventHandler MouseMove;

        /// <summary>
        /// Occurs when the user releases the mouse button while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event ToolsMouseEventHandler MouseUp;

        #endregion
        
        #region Properties

        /// <summary>
        /// Tray Item Menu
        /// </summary>
        public TrayMenu Menu { get; private set; }

        /// <summary>
        /// Dont show if the owner application is in automation mode
        /// </summary>
        public bool SuspendOnAutomation
        {
            get
            {
                return _suspendOnAutomation;
            }
            set
            {
                if (value != _suspendOnAutomation)
                {
                    _suspendOnAutomation = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the ToolTip text displayed when the mouse pointer rests on a notification area icon
        /// </summary>
        public string Text
        {
            get
            {
                return null != _icon ? _icon.Text : String.Empty;
            }
            set
            {
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.Text = value;
            }
        }

        /// <summary>
        /// Gets or sets the title of the balloon tip displayed on the NotifyIcon
        /// </summary>
        public string BalloonTipTitle
        {
            get
            {
                return null != _icon ? _icon.BalloonTipTitle : String.Empty;
            }
            set
            {
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.BalloonTipTitle = value;
            }
        }

        /// <summary>
        /// Gets or sets the text to display on the balloon tip associated with the NotifyIcon
        /// </summary>
        public string BallonTipText
        {
            get 
            {
                return null != _icon ? _icon.BalloonTipText : String.Empty;
            }
            set
            {
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.BalloonTipText = value;
            }
        }

        /// <summary>
        /// Gets or sets the icon to display on the balloon tip associated with the NotifyIcon
        /// </summary>
        public TrayToolTipIcon BallonTipIcon
        {
            get
            {
                return null != _icon ? (TrayToolTipIcon)_icon.BalloonTipIcon : TrayToolTipIcon.None;
            }
            set
            {
                if (TrayToolTipIcon.None == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.BalloonTipIcon = (ToolTipIcon)value;
            }
        }
         
        /// <summary>
        /// Gets or sets the current icon
        /// </summary>
        public Icon Icon
        {
            get
            {
                return null != _icon ? _icon.Icon : ApplicationIcon;
            }
            set
            {
                if (null != _icon)
                    _icon.Icon = value;
                else
                    _icon = CreateConnectTray(value);
            }
        }
         
        /// <summary>
        /// Gets or sets a value indicating whether the icon is visible in the notification area of the taskbar
        /// </summary>
        public bool Visible
        {
            get
            {
                return null != _icon ? _visible : false;
            }
            set
            {
                _visible = value;
                if (null == _icon && false == value)
                    return;

                if (value)
                {
                    if(null == _icon)
                        _icon = CreateConnectTray();
                }
                else
                {
                    if (null != _icon)
                    { 
                        DisposeTray();
                        _icon = null;
                    }
                }

                if (null != _icon)
                    _icon.Visible = GetEffectiveVisibility();
            }
        }

        /// <summary>
        /// Application Icon in executing assembly
        /// </summary>
        protected virtual Icon ApplicationIcon
        {
            get
            {
                return TryGetApplicationIcon();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Setup tray as one call
        /// </summary>
        /// <param name="visible">tray visibility</param>
        /// <param name="text">shown text</param>
        public void Setup(bool visible, string text)
        {
            Visible = visible;
            Text = text;
        }

        /// <summary>
        /// Setup tray as one call
        /// </summary>
        /// <param name="visible">tray visibility</param>
        /// <param name="text">shown text</param>
        /// <param name="icon">shown icon</param>
        public void Setup(bool visible, string text, Icon icon)
        {
            Visible = visible;
            Text = text;
            Icon = icon;
        }

        /// <summary>
        /// Setup tray as one call
        /// </summary>
        /// <param name="visible">tray visibility</param>
        /// <param name="text">shown text</param>
        /// <param name="iconResource">full qualified icon resource address</param>
        public void Setup(bool visible, string text, string iconResource)
        {
            Visible = visible;
            Text = text;
            System.IO.Stream iconStream = ReadRessource(iconResource);
            if (null != iconStream)
                Icon = new System.Drawing.Icon(iconStream);
        }

        /// <summary>
        /// Displays a balloon tip in the taskbar for the specified time period
        /// </summary>
        /// <param name="timeout">The time period, in milliseconds, the balloon tip should display</param>
        public void ShowBalloonTip(int timeout)
        {
            if (null == _icon)
                _icon = CreateConnectTray();
            _icon.ShowBalloonTip(timeout);
        }

        /// <summary>
        /// Displays a balloon tip with the specified title, text, and icon in the taskbar for the specified time period
        /// </summary>
        /// <param name="timeout">The time period, in milliseconds, the balloon tip should display</param>
        /// <param name="tipTitle">The title to display on the balloon tip.</param>
        /// <param name="tiptext">The text to display on the balloon tip</param>
        /// <param name="tipIcon">One of the ToolTipIcon values</param>
        public void ShowBalloonTip(int timeout, string tipTitle, string tiptext, TrayToolTipIcon tipIcon)
        {
            if (null == _icon)
                _icon = CreateConnectTray();
            _icon.ShowBalloonTip(timeout, tipTitle, tiptext, (ToolTipIcon)tipIcon);
        }

        /// <summary>
        /// Bring up the menu if exists
        /// </summary>
        internal void ShowContextMenu()
        {
            if (null != _icon.ContextMenuStrip)
            {
                MethodInfo showMethod = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                showMethod.Invoke(_icon, null);
            }
        }

        /// <summary>
        /// Creates a container as owner for new NotifyIcon instances
        /// </summary>
        /// <returns>IContainer instance or null</returns>
        protected internal virtual IContainer CreateContainer()
        {
            return null;
        }
      
        /// <summary>
        /// Creates NotifyIcon for the instance
        /// </summary>
        /// <returns>inner NotifyIcon instance</returns>
        protected internal virtual NotifyIcon CreateTray()
        {
            NotifyIcon result = null;
            IContainer container = CreateContainer();
            if (null != container)
                result = new NotifyIcon(container);
            else
                result = new NotifyIcon();

            result.ContextMenuStrip = Menu.GetMenuInternal<ContextMenuStrip>();

            result.Visible = GetEffectiveVisibility();
            result.Icon = null != Icon ? Icon : ApplicationIcon;
            return result;
        }

        /// <summary>
        /// Dispose inner NotifyIcon instance
        /// </summary>
        protected internal virtual void DisposeTrayIcon()
        {
            if (null != _icon)
            {                
                _icon.Dispose();
                _icon = null;
            }
        }

        /// <summary>
        /// Dispose inner menu instance
        /// </summary>
        protected internal virtual void DisposeMenu()
        {
            if (null != Menu && false == Menu.IsDisposed)
                Menu.Dispose();
        }

        /// <summary>
        /// Dispose inner NotifyIcon instance
        /// </summary>
        internal void DisposeTray()
        {
            DisconnectEvents(_icon);
            DisposeMenu();
            DisposeTrayIcon();           
        }

        private NotifyIcon CreateConnectTray(Icon value)
        {
            NotifyIcon icon = CreateTray();
            icon.Icon = value;
            return icon;
        }

        private NotifyIcon CreateConnectTray()
        {
            NotifyIcon icon = CreateTray();
            if (null != icon)
                ConnectEvents(icon);
            return icon;
        }

        private bool GetEffectiveVisibility()
        {
            if (SuspendOnAutomation && _owner.IsAutomation)
                return false;
            else
                return _visible;
        }

        private void ConnectEvents(NotifyIcon icon)
        {
            if (null == icon)
                return;

            icon.BalloonTipClicked += new EventHandler(Icon_BalloonTipClicked);
            icon.BalloonTipClosed += new EventHandler(Icon_BalloonTipClosed);
            icon.BalloonTipShown += new EventHandler(Icon_BalloonTipShown);
            icon.Click += new EventHandler(Icon_Click);
            icon.DoubleClick += new EventHandler(Icon_DoubleClick);
            icon.MouseClick += new MouseEventHandler(Icon_MouseClick);
            icon.MouseDoubleClick += new MouseEventHandler(Icon_MouseDoubleClick);
            icon.MouseDown += new MouseEventHandler(Icon_MouseDown);
            icon.MouseMove += new MouseEventHandler(Icon_MouseMove);
            icon.MouseUp += new MouseEventHandler(Icon_MouseUp);
        }

        private void DisconnectEvents(NotifyIcon icon)
        {
            if (null == icon)
                return;

            icon.BalloonTipClicked -= Icon_BalloonTipClicked;
            icon.BalloonTipClosed -= Icon_BalloonTipClosed;
            icon.BalloonTipShown -= Icon_BalloonTipShown;
            icon.Click -= Icon_Click;
            icon.DoubleClick -= Icon_DoubleClick;
            icon.MouseClick -= Icon_MouseClick;
            icon.MouseDoubleClick -= Icon_MouseDoubleClick;
            icon.MouseDown -= Icon_MouseDown;
            icon.MouseMove -= Icon_MouseMove;
            icon.MouseUp -= Icon_MouseUp;
        }

        private Icon TryGetApplicationIcon()
        {
            try
            {
                if (_applicationIconResolved)
                    return _applicationIcon;

                _applicationIcon = Icon.ExtractAssociatedIcon(_owner.OwnerAssembly.Location);

                return _applicationIcon;
            }
            catch
            {
                return null;
            }
            finally
            {
                _applicationIconResolved = true;
            }
        }

        private System.IO.Stream ReadRessource(string address)
        {
            if (null == _owner || null == _owner.Owner || null == _owner.OwnerAssembly)
                return null;

            System.Reflection.Assembly assembly = _owner.OwnerAssembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream(address);
            if (null == stream)
            {
                string space = _owner.Owner.GetType().Namespace;
                stream = assembly.GetManifestResourceStream(space + "." + address);
            }
            return stream;
        }

        #endregion

        #region Trigger

        private void Icon_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (null != MouseUp)
                    MouseUp(this, new ToolsMouseEventArgs((ToolsMouseButtons)e.Button, e.Clicks, e.X, e.Y, e.Delta));
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {                 
                if (null != MouseMove)
                    MouseMove(this, new ToolsMouseEventArgs((ToolsMouseButtons)e.Button, e.Clicks, e.X, e.Y, e.Delta));
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (null != MouseDown)
                    MouseDown(this, new ToolsMouseEventArgs((ToolsMouseButtons)e.Button, e.Clicks, e.X, e.Y, e.Delta));
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (null != MouseDoubleClick)
                    MouseDoubleClick(this, new ToolsMouseEventArgs((ToolsMouseButtons)e.Button, e.Clicks, e.X, e.Y, e.Delta));
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (null != MouseClick)
                    MouseClick(this, new ToolsMouseEventArgs((ToolsMouseButtons)e.Button, e.Clicks, e.X, e.Y, e.Delta));
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (null != DoubleClick)
                    DoubleClick(this, e);
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_Click(object sender, EventArgs e)
        {
            try
            {
                MouseEventArgs args = e as MouseEventArgs;
                if (true == Menu.Enabled && args.Button == MouseButtons.Left && Menu.ClickMode == TrayMenuClickMode.LeftRight)
                    ShowContextMenu();

                if (null != Click)
                    Click(this, e);
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_BalloonTipShown(object sender, EventArgs e)
        {
            try
            {
                if (null != BalloonTipShown)
                    BalloonTipShown(this, e);
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_BalloonTipClosed(object sender, EventArgs e)
        {
            try
            {
                if (null != BalloonTipClosed)
                    BalloonTipClosed(this, e);
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        private void Icon_BalloonTipClicked(object sender, EventArgs e)
        {
            try
            {
                if (null != BalloonTipClicked)
                    BalloonTipClicked(this, e);
            }
            catch (Exception exception)
            {
                _owner.OwnerApplication.Console.WriteException(exception);
                throw;
            }
        }

        #endregion
    }
}