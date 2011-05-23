using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace NetOffice.DeveloperUtils
{
    interface IUtilsControl
    {
        string ControlName { get; }

        /// <summary>
        /// Control visible 
        /// </summary>
        void Activate();

        void LoadConfiguration(XmlNode configNode);

        void SaveConfiguration(XmlNode configNode);
        
        /// <summary>
        /// substitute for dispose
        /// </summary>
        void Release();
    }
}
