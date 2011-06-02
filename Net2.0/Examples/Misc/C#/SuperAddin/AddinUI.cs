using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
 
namespace SuperAddin
{
    /// <summary>
    /// represents the addin user interace
    /// </summary>
    internal class AddinUI
    {
        #region Fields

        HostApplication _application;
        bool            _ribbonActive;

        #endregion

        #region Construction

        internal AddinUI(HostApplication application)
        {
            _application = application;
        }
        
        #endregion

        #region Properties

        public HostApplication Application
        {
            get 
            {
                return _application;
            }
        }

        public RibbonUI RibbonUI
        {
            get
            {
                return _ribbonUI;
            }
        }

        public ClassicUI ClassicUI
        {
            get
            {
                return _classicUI;
            }
        }

        public bool RibbonIsActive
        {
            get
            {
                return _ribbonActive;
            }
            internal set
            {
                _ribbonActive = value;
            }
        }
        
        #endregion

        #region Events

        public event ButtonClickEventHandler ButtonClick;

        #endregion

        #region Event Hell Raiser

        internal void RaiseButtonClick(ButtonClickArgs clickArgs)
        {
            if (null != ButtonClick)
                ButtonClick(clickArgs);
            else
                clickArgs.Dispose();
        }

        #endregion
    }
}
