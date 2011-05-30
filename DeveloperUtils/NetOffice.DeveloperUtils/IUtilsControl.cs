using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace NetOffice.DeveloperUtils
{
    interface IUtilsControl
    {
        /// <summary>
        /// returns the name of control, displayed in application tab
        /// </summary>
        string ControlName { get; }

        /// <summary>
        /// method was called from host application while application tab selection is changed to control 
        /// </summary>
        void Activate();

        /// <summary>
        ///  method was called from host application after start
        /// </summary>
        /// <param name="configNode"></param>
        void LoadConfiguration(XmlNode configNode);

        /// <summary>
        /// method was called from host application before close
        /// </summary>
        /// <param name="configNode"></param>
        void SaveConfiguration(XmlNode configNode);

        /// <summary>
        /// method was called from after start and after selection change from user
        /// 0 = english, 1 = german
        /// </summary>
        /// <param name="id"></param>
        void SetLanguage(int id);

        /// <summary>
        /// method was called from host application before close, substitute for dispose
        /// </summary>
        void Release();
    }
}
