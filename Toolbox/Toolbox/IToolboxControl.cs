using System;
using System.IO;
using System.Xml;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// toolboxcontrol message kind. the application is showing a corresponding icon.
    /// </summary>
    public enum ToolboxControlMessageKind
    { 
        /// <summary>
        /// message is an information
        /// </summary>
        Information = 0,
        
        /// <summary>
        /// message is a warning
        /// </summary>
        Warning = 1,

        /// <summary>
        /// message is uncategorized 
        /// </summary>
        Uncategorized
    }

    /// <summary>
    /// Represents a toolbox content control
    /// </summary>
    public interface IToolboxControl : IDisposable 
    {
        /// <summary>
        /// parent host application
        /// </summary>
        IToolboxHost Host { get; }

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
        /// returns the instance supports help text content
        /// </summary>
        bool SupportsHelpContent { get; }

        /// <summary>
        /// returns info the instance want show a message in the upper area
        /// </summary>
        bool SupportsInfoMessage { get; }

        /// <summary>
        /// info message kind
        /// </summary>
        ToolboxControlMessageKind InfoMessageKind { get; }

        /// <summary>
        /// additional message displayed in the upper area
        /// </summary>
        string InfoMessage { get; }

        /// <summary>
        /// initialize the instance
        /// </summary>
        /// <param name="host">host application</param>
        void InitializeControl(IToolboxHost host);

        /// <summary>
        /// method was called from host application while application tab selection is changed to control 
        /// </summary>
        /// <param name="firstTime">control is shown first time</param>
        void Activate(bool firstTime);

        /// <summary>
        /// control is not visible any longer because user switch to another tabpage
        /// </summary>
        void Deactivated();

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
        /// returns help richtext if supported, otherwise a NotImplementedException is thrown
        /// </summary>
        /// <param name="lcid">target language id</param>
        /// <returns>help content as rich text(.rtf)</returns>
        Stream GetHelpText();

        /// <summary>
        /// redirected from host application if control is currently visible
        /// </summary>
        /// <param name="e"></param>
        void KeyDown(KeyEventArgs e);

        /// <summary>
        /// custom instance destructor
        /// </summary>
        void Release();
    }
}