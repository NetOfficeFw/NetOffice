using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace SuperAddinCSharp
{
    public class TrayIcon : IDisposable
    {
        #region Fields

        NotifyIcon _trayIcon;
        
        #endregion

        #region Construction

        public TrayIcon(bool visible)
        {
            _trayIcon = new NotifyIcon(new System.ComponentModel.Container());
            System.IO.Stream iconStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("SuperAddin.Properties.AddinIcon.ico");
            _trayIcon.Icon = new System.Drawing.Icon(iconStream);
            iconStream.Close();
            _trayIcon.Text = "SuperAdddin loaded.";
            _trayIcon.Visible = visible;
        }

        #endregion
        
        #region Properties
        
        public bool Visible
        {
            get
            {
                return _trayIcon.Visible;
            }
            set
            {
                _trayIcon.Visible = value;
            }
        }
        
        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (null != _trayIcon)
                _trayIcon.Dispose();
        }

        #endregion
    }
}
