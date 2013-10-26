using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Manage the application tray icon
    /// </summary>
    internal class TrayIconManager : IDisposable
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public TrayIconManager(FormMain parent, ContextMenuStrip trayStrip)
        {
            Parent = parent;
            InnerIcon = new NotifyIcon();
            InnerIcon.Visible = true;
            InnerIcon.Text = GetType().Assembly.GetName().Name;
            InnerIcon.Icon = NormalIcon;
            InnerIcon.MouseClick += new MouseEventHandler(InnerIcon_MouseClick);
            InnerIcon.ContextMenuStrip = trayStrip;
        }

        #endregion

        #region Properties

        private FormMain Parent { get; set; }

        private NotifyIcon InnerIcon { get; set; }

        private Icon NormalIcon 
        {
            get
            {
                if (null == _normalIcon)
                    _normalIcon = GetIconFromRessource("Application.ico");
                return _normalIcon;        
            }
        }
        private Icon _normalIcon;

        private Icon UpdateIcon
        {
            get
            {
                if (null == _updateIcon)
                    _updateIcon = GetIconFromRessource("ApplicationInfo.ico");
                return _updateIcon;
            }
        }
        private Icon _updateIcon;

        private object LockThis
        {
            get 
            {
                if (null == _lockThis)
                    _lockThis = new object();
                return _lockThis;
            }
        }
        private object _lockThis;

        private Timer UpdateTimer
        {
            get
            {
                if (null == _updateTimer)
                {
                    _updateTimer = new Timer(new System.ComponentModel.Container());
                    _updateTimer.Interval = 1000;
                    _updateTimer.Tick += new EventHandler(UpdateTimer_Tick);
                }
                return _updateTimer;
            }
        }
        private Timer _updateTimer;

        private DateTime? StartUpdateTime{get;set;}

        #endregion

        #region IDisposable Member

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            if (null != InnerIcon)
            { 
                InnerIcon.Dispose();
                InnerIcon = null;                 
            }

            if (null != _normalIcon)
            {
                _normalIcon.Dispose();
                _normalIcon = null;
            }

            if (null != _updateIcon)
            {
                _updateIcon.Dispose();
                _updateIcon = null;
            }

        }

        #endregion

        #region Methods
        
        /// <summary>
        /// Shows a different trayicon for 1 second
        /// </summary>
        internal void ShowUpdate()
        {
            lock (LockThis)
            {
                if (null == StartUpdateTime)
                    InnerIcon.Icon = UpdateIcon;
                StartUpdateTime = DateTime.Now;
                UpdateTimer.Enabled = true;
            }
        }

        private void StopUpdate()
        {
            lock (LockThis)
            {
                UpdateTimer.Enabled = false;
                StartUpdateTime = null;
                InnerIcon.Icon = NormalIcon;
            }        
        }

        private static Icon GetIconFromRessource(string name)
        {
            Stream resssourceStream = Assembly.GetCallingAssembly().GetManifestResourceStream(string.Format("{0}.Icons.{1}", Assembly.GetCallingAssembly().GetName().Name, name));
            if (null == resssourceStream)
                throw new System.IO.FileLoadException(name);
            resssourceStream.Seek(0, SeekOrigin.Begin);
            Icon ressourceIcon = new Icon(resssourceStream);
            return ressourceIcon;
        }

        #endregion

        #region Trigger

        private void InnerIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && null != Parent)
            {
                Parent.Show();
                Parent.WindowState = FormWindowState.Normal;
            }
        }
         
        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            if (null == StartUpdateTime)
                return;
            
            DateTime elapseTimeLimit = StartUpdateTime.Value.AddSeconds(1);
            DateTime now = DateTime.Now;

            if (now >= elapseTimeLimit)
                StopUpdate();
        }

        #endregion
    }
}
