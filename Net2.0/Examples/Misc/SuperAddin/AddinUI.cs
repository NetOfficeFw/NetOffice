using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using SuperAddin.UIMapper;
 
namespace SuperAddin
{
    public class ButtonClickArgs : EventArgs , IDisposable
    {
        #region Fields

        private IRibbonControl _ribbonControl;
        private LateBindingApi.OfficeApi.CommandBarButton _buttonControl;
        
        #endregion

        #region Construction
        
        internal ButtonClickArgs(IRibbonControl ribbonControl)
        {
            _ribbonControl = ribbonControl;
        }

        internal ButtonClickArgs(LateBindingApi.OfficeApi.CommandBarButton buttonControl)
        {
            _buttonControl = buttonControl;
        }

        #endregion

        #region Properties

        public IRibbonControl RibbonControl
        {
            get 
            {
                return _ribbonControl;
            }
        }

        public LateBindingApi.OfficeApi.CommandBarButton ButtonControl
        {
            get 
            {
                return _buttonControl;
            }
        }
        
        #endregion

        #region IDisposable Members
        
        public void Dispose()
        {
            if (null != _ribbonControl)
                Marshal.ReleaseComObject(_ribbonControl);

            if (null != _buttonControl)
                _buttonControl.Dispose();
        }

        #endregion
    }

    public delegate void ButtonClickEventHandler(ButtonClickArgs args);

    /// <summary>
    /// represents the addin user interace
    /// </summary>
    internal class AddinUI
    {
        #region Fields

        HostApplication _application;
        RibbonUI        _ribbonUI;
        ClassicUI       _classicUI;
        bool            _ribbonActive;

        #endregion

        #region Construction

        internal AddinUI(HostApplication application)
        {
            _ribbonUI = new UIMapper.RibbonUI(this);
            _classicUI = new UIMapper.ClassicUI(this);
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
