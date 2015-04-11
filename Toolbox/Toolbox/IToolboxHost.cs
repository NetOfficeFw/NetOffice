using System;
using System.Drawing;
using NetOffice.DeveloperToolbox.Translation;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Represents the host application
    /// </summary>
    public interface IToolboxHost
    {
        /// <summary>
        /// Application Name
        /// </summary>
        string Caption { get; }

        /// <summary>
        /// Available Application Languages
        /// </summary>
        ToolLanguages Languages { get; }

        /// <summary>
        /// The host is supporting a language editor
        /// </summary>
        bool SupportsLanguageEditor { get; }

        /// <summary>
        /// Get or set language editor visibilty
        /// </summary>
        bool LanguageEditorVisible { get; set; }

        // Occurs when the langage editor visibilty has been changed
        event EventHandler LanguageEditorVisibleChanged;

        /// <summary>
        /// Current Language ID
        /// </summary>
        int CurrentLanguageID { get; set; }

        /// <summary>
        /// Application Icon
        /// </summary>
        Icon Icon { get; }

        /// <summary>
        /// Current loaded content controls
        /// </summary>
        IToolboxControl[] ToolboxControls { get; }

        /// <summary>
        /// Switch to toolbox control
        /// </summary>
        /// <param name="controlName">logical name of the control</param>
        void SwitchTo(string controlName);

        /// <summary>
        /// Display the application main window
        /// </summary>
        void ShowMainWindow();

        /// <summary>
        /// Minimize the application main window
        /// </summary>
        /// <param name="showInTaskbar">true if visible in taskbar, otherwise false</param>
        void MinimizeMainWindow(bool showInTaskbar);

        /// <summary>
        /// Occurs when the application main window has been minimized
        /// </summary>
        event EventHandler Minimized;
    }
}
