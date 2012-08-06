using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    interface IToolboxControl : IDisposable 
    {
        /// <summary>
        /// returns the name of control
        /// </summary>
        string ControlName { get; }

        /// <summary>
        /// returns the caption of control, displayed in application tab
        /// </summary>
        string ControlCaption { get; }

        /// <summary>
        /// returns the icon of control, displayed in application tab
        /// </summary>
        Image Icon { get; }

        /// <summary>
        /// method was called from host application while application tab selection is changed to control 
        /// </summary>
        void Activate();
       
        /// <summary>
        /// method was called when application is completly loaded
        /// </summary>
        void LoadComplete();

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
        /// 1031 = english, 1033 = german
        /// </summary>
        /// <param name="id"></param>
        void SetLanguage(int id);

        /// <summary>
        /// redirected from host application
        /// </summary>
        /// <param name="e"></param>
        void KeyDown(KeyEventArgs e);

        /// <summary>
        /// components from control, translate stuff
        /// </summary>
        System.ComponentModel.IContainer Components { get; }
    }
}
