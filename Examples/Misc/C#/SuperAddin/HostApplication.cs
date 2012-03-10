using System;
using Extensibility;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;

using LateBindingApi.Core;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;

namespace SuperAddinCSharp
{
    /// <summary>
    /// represents the office host application
    /// </summary>
    public class HostApplication : IDisposable
    {
        #region Fields

        COMObject                 _hostApp;

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
        public string ComponentName
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
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (null != _hostApp)
                _hostApp.Dispose();
        }

        #endregion
    }
}
