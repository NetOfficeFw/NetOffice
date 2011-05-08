using System;
using Extensibility;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;

using LateBindingApi.Core;
using Office = LateBindingApi.OfficeApi;
using Excel = LateBindingApi.ExcelApi;
using Word = LateBindingApi.WordApi;
using Outlook = LateBindingApi.OutlookApi;
using PowerPoint = LateBindingApi.PowerPointApi;
using Access = LateBindingApi.AccessApi;

namespace SuperAddin
{
    /// <summary>
    /// represents the office host application
    /// </summary>
    public class HostApplication : IDisposable
    {
        #region Fields

        COMObject                 _hostApp;
        DocumentEventsMapper      _docEvents;

        #endregion

        #region Events

        public event BeforeSaveHandler  BeforeSave;
        public event OpenHandler        BeforeOpen;
        public event BeforeCloseHandler BeforeClose;
        public event BeforePrintHandler BeforePrint;
        
        #endregion

        #region Properties

        /// <summary>
        /// gets office host application object
        /// </summary>
        public COMObject Application
        {
            get 
            {
                return _hostApp;
            }
        }
        
        /// <summary>
        /// returns component host name of host app
        /// </summary>
        public string Component
        {
            get
            {
                return TypeDescriptor.GetComponentName(_hostApp.UnderlyingObject);
            }
        }

        /// <summary>
        /// returns name of host app
        /// </summary>
        public string Name
        {
            get
            {
               return _hostApp.UnderlyingTypeName;
            }
        }

        /// <summary>
        /// returns version of host app
        /// </summary>
        public string Version
        {
            get 
            {
                return (string)Invoker.PropertyGet(_hostApp.UnderlyingObject, "Version");
            }
        }

        #endregion

        #region Construction

        public HostApplication(object comProxy, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            LateBindingApi.Core.Factory.Initialize();
            LateBindingApi.Core.Settings.EnableEvents = true;

            string typeComponent = System.ComponentModel.TypeDescriptor.GetComponentName(comProxy);
            switch (typeComponent)
            { 
                case "Microsoft Excel":
                case "Excel":
                    _hostApp = new Excel.Application(null, comProxy);
                    break;
                case "Microsoft Word":
                case "Word":
                    _hostApp = new Word.Application(null, comProxy);
                    break;
                case "Microsoft Outlook":
                case "Outlook":
                    _hostApp = new Outlook.Application(null, comProxy);
                    break;
                case "Microsoft PowerPoint":
                case "PowerPoint":
                    _hostApp = new PowerPoint.Application(null, comProxy);
                    break;
                case "Microsoft Access":
                case "Access":
                    _hostApp = new Access.Application(null, comProxy);
                    break;
            }

            _docEvents = new DocumentEventsMapper(this);
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (null != _hostApp)
                _hostApp.Dispose();
        }

        #endregion

        #region Event Hell Raiser

        internal void RaiseBeforeSaveEvent(BeforeSaveArgs args, ref bool SaveAsUI, ref bool Cancel)
        {
            if (null != BeforeSave)
                BeforeSave(args, ref SaveAsUI, ref Cancel);
            else
                args.Dispose();
        }

        internal void RaiseBeforeOpenEvent(OpenArgs args)
        {
            if (null != BeforeOpen)
                BeforeOpen(args);
            else
                args.Dispose();
        }

        internal void RaiseBeforeCloseEvent(BeforeCloseArgs args, ref bool Cancel)
        {
            if (null != args)
                BeforeClose(args, ref Cancel);
            else
                args.Dispose();
        }

        internal void RaiseBeforePrintEvent(BeforePrintArgs args, ref bool Cancel)
        {
            if (null != args)
                BeforePrint(args, ref Cancel);
            else
                args.Dispose();
        }

        #endregion
    }
}
