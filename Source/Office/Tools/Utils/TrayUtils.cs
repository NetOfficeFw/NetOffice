using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Text;

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
        public event MouseEventHandler MouseClick;
        
        /// <summary>
        /// Occurs when the user double-clicks the NotifyIcon with the mouse
        /// </summary>
        public event MouseEventHandler MouseDoubleClick;
        
        /// <summary>
        /// Occurs when the user presses the mouse button while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event MouseEventHandler MouseDown;
        
        /// <summary>
        /// Occurs when the user moves the mouse while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event MouseEventHandler MouseMove;

        /// <summary>
        /// Occurs when the user releases the mouse button while the pointer is over the icon in the notification area of the taskbar
        /// </summary>
        public event MouseEventHandler MouseUp;

        #endregion

        #region Properties

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
        public ToolTipIcon BallonTipIcon
        {
            get
            {
                return null != _icon ? _icon.BalloonTipIcon : ToolTipIcon.None;
            }
            set
            {
                if (ToolTipIcon.None == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.BalloonTipIcon = value;
            }
        }
       
        /// <summary>
        /// Gets or sets the shortcut menu associated with the NotifyIcon
        /// </summary>
        public ContextMenuStrip ContextMenuStrip
        {
            get
            {
                return null != _icon ? _icon.ContextMenuStrip : null;
            }
            set
            {
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.ContextMenuStrip = value;
            }
        }

        /// <summary>
        /// Gets or sets the shortcut menu for the icon
        /// </summary>
        public ContextMenu ContextMenu
        {
            get 
            {
                return null != _icon ? _icon.ContextMenu : null;
            }
            set
            {
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.ContextMenu = value;
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
                if (null == value && null == _icon)
                    return;
                if (null == _icon)
                    _icon = CreateConnectTray();
                _icon.Icon = value;
            }
        }
        
        /// <summary>
        /// Gets or sets a value indicating whether the icon is visible in the notification area of the taskbar
        /// </summary>
        public bool Visible
        {
            get
            {
                return null != _icon && _icon.Visible == true;
            }
            set
            {
                if (null == _icon && false == value)
                    return;

                if (value)
                {
                    if(null == _icon)
                        _icon = CreateConnectTray();
                    _icon.Visible = value;
                }
                else
                {
                    if (null != _icon)
                    { 
                        DisposeTray();
                        _icon = null;
                    }
                }
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
        /// <param name="contextMenu">shortcut menu</param>
        /// <param name="icon">shown icon</param>
        public void Setup(bool visible, string text, ContextMenu contextMenu, Icon icon)
        {
            Visible = visible;
            Text = text;
            ContextMenu = contextMenu;
            Icon = icon;
        }

        public void Setup(bool visible, string text, ContextMenu contextMenu, string iconResource)
        {
            throw new NotImplementedException();
        }

        public void Setup(bool visible, string text, string iconResource)
        {
            throw new NotImplementedException();
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
        public void ShowBalloonTip(int timeout, string tipTitle, string tiptext, ToolTipIcon tipIcon)
        {
            if (null == _icon)
                _icon = CreateConnectTray();
            _icon.ShowBalloonTip(timeout, tipTitle, tiptext, tipIcon);
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

            result.Icon = ApplicationIcon;
            return result;
        }

        /// <summary>
        /// Dispose inner NotifyIcon instance
        /// </summary>
        /// <param name="icon">inner NotifyIcon instance</param>
        protected internal virtual void DisposeTray(NotifyIcon icon)
        {
            if (null != icon)
                icon.Dispose();
        }

        /// <summary>
        /// Dispose inner NotifyIcon instance
        /// </summary>
        internal void DisposeTray()
        {
            DisconnectEvents(_icon);
            DisposeTray(_icon);
            _icon = null;
        }

        private NotifyIcon CreateConnectTray()
        {
            NotifyIcon icon = CreateTray();
            if (null != icon)
                ConnectEvents(icon);
            return icon;
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

        #endregion

        #region Trigger

        private void Icon_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (null != MouseUp)
                    MouseUp(this, e);
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
                    MouseMove(this, e);
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
                    MouseDown(this, e);
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
                    MouseDoubleClick(this, e);
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
                    MouseClick(this, e);
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