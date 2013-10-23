using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;

namespace NOTools.CodeCommander.Logic
{
    /// <summary>
    /// handles the different possible office host applications
    /// </summary>
    internal class OfficeApplicationManager
    {
        /// <summary>
        /// creates an instance of the class
        /// </summary>
        /// <param name="parent">office host application</param>
        internal OfficeApplicationManager(COMObject parent)
        {
            Parent = parent;
        }

        /// <summary>
        /// Office Host Application
        /// </summary>
        internal COMObject Parent { get; set; }

        public COMObject GetSelectedProxy(int index)
        {
            string parentTypeName = Parent.GetType().FullName;
            switch (parentTypeName)
            {
                case "NetOffice.ExcelApi.Application":
                    return GetSelectedExcelProxy(index);
                case "NetOffice.WordApi.Application":
                    return GetSelectedWordProxy(index);
                case "NetOffice.OutlookApi.Application":
                    return GetSelectedOutlookProxy(index);
                case "NetOffice.PowerPointApi.Application":
                    return GetSelectedPowerPointProxy(index);
                case "NetOffice.AccessApi.Application":
                    return GetSelectedAccessProxy(index);
                case "NetOffice.MSProjectApi.Application":
                    return GetSelectedProjectProxy(index);
                default:
                    return null;
            }
        }

        private COMObject GetSelectedExcelProxy(int index)
        {
            switch(index)
            {
                case 0:
                    return Parent;
                case 1:
                    return (Parent as NetOffice.ExcelApi.Application).ActiveWorkbook;
                case 2:
                    return (Parent as NetOffice.ExcelApi.Application).ActiveSheet as NetOffice.ExcelApi.Worksheet;
            }

            return null;
        }

        private COMObject GetSelectedWordProxy(int index)
        {
            return null;
        }

        private COMObject GetSelectedOutlookProxy(int index)
        {
            return null;
        }

        private COMObject GetSelectedPowerPointProxy(int index)
        {
            return null;
        }

        private COMObject GetSelectedAccessProxy(int index)
        {
            return null;
        }

        private COMObject GetSelectedProjectProxy(int index)
        {
            return null;
        }

        /// <summary>
        /// returns the available proxies for the propertygrid depending on the current office application
        /// </summary>
        /// <returns></returns>
        public AvailableProxy[] GetAvailableProxies()
        {
            string parentTypeName = Parent.GetType().FullName;
            switch (parentTypeName)
            {
                case "NetOffice.ExcelApi.Application":
                    return GetAvailableExcelProxies();
                case "NetOffice.WordApi.Application":
                    return GetAvailableWordProxies();
                case "NetOffice.OutlookApi.Application":
                    return GetAvailableOutlookProxies();
                case "NetOffice.PowerPointApi.Application":
                    return GetAvailablePowerPointProxies();
                case "NetOffice.AccessApi.Application":
                    return GetAvailableAccessProxies();
                case "NetOffice.MSProjectApi.Application":
                    return GetAvailableProjectProxies();
                default:
                    return new AvailableProxy[0];
            }
        }

        private AvailableProxy[] GetAvailableExcelProxies()
        {
            List<AvailableProxy> list = new List<AvailableProxy>();
            list.Add(new AvailableProxy(0, "Application"));
            list.Add(new AvailableProxy(1, "Current Workbook"));
            list.Add(new AvailableProxy(2, "Current Worksheet"));
            return list.ToArray();
        }

        private AvailableProxy[] GetAvailableWordProxies()
        {
            return null;
        }

        private AvailableProxy[] GetAvailableOutlookProxies()
        {
            return null;
        }

        private AvailableProxy[] GetAvailablePowerPointProxies()
        {
            return null;
        }

        private AvailableProxy[] GetAvailableAccessProxies()
        {
            return null;
        }

        private AvailableProxy[] GetAvailableProjectProxies()
        {
            return null;
        }
    }
}
