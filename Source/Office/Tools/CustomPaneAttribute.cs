using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Pane Dock Position
    ///  SupportByVersion Office 12, 14, 15, 16
    /// </summary>
    public enum PaneDockPosition
    {
        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionLeft = 0,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionTop = 1,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionRight = 2,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionBottom = 3,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionFloating = 4
    }

    /// <summary>
    /// Pane Dock Restrictions
    /// SupportByVersion Office 12, 14, 15, 16
    /// </summary>
    public enum PaneDockPositionRestrict
    {
        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionRestrictNone = 0,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionRestrictNoChange = 1,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionRestrictNoHorizontal = 2,

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        msoCTPDockPositionRestrictNoVertical = 3
    }

    /// <summary>
    ///
    /// SupportByVersion Office 12, 14, 15, 16
    /// </summary>
    public enum PaneCreation
    {
        /// <summary>
        /// Create pane automatically while addin startup
        /// </summary>
        AutomaticallyAtStartup = 0,

        /// <summary>
        /// Create pane at hand using the TaskpaneInfo->Create method.
        /// Attributed panes are available in the addin instance property Taskpanes.
        /// </summary>
        Manually = 1
    }

    /// <summary>
    /// Specify a custom task pane
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = true)]
    public class CustomPaneAttribute : System.Attribute
    {
        #region Fields

        /// <summary>
        /// Type of the custom task pane
        /// </summary>
        public readonly Type PaneType;

        /// <summary>
        /// Pane Title (Default is Pane Type Name)
        /// </summary>
        public readonly string Title;

        /// <summary>
        /// Pane Visibility (Default is true)
        /// </summary>
        public readonly bool Visible;

        /// <summary>
        /// Pane dock alignment direction (Default is Right)
        /// </summary>
        public readonly PaneDockPosition DockPosition;

        /// <summary>
        /// Pane dock alignment restriction (Default is None)
        /// </summary>
        public readonly PaneDockPositionRestrict DockPositionRestrict;

        /// <summary>
        /// Pane Width (Default is 150)
        /// </summary>
        public readonly int Width;

        /// <summary>
        /// Pane Height (Default is 150)
        /// </summary>
        public readonly int Height;

        /// <summary>
        /// Pane Creation Mode
        /// </summary>
        public readonly PaneCreation Creation;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        public CustomPaneAttribute(Type paneType)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = paneType.Name;
            Visible = true;
            DockPosition = PaneDockPosition.msoCTPDockPositionRight;
            DockPositionRestrict = PaneDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        public CustomPaneAttribute(Type paneType, string title)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Visible = true;
            DockPosition = PaneDockPosition.msoCTPDockPositionRight;
            DockPositionRestrict = PaneDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = PaneDockPosition.msoCTPDockPositionRight;
            DockPositionRestrict = PaneDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="creation">pane creation mode</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneCreation creation)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            Creation = creation;
            DockPosition = PaneDockPosition.msoCTPDockPositionRight;
            DockPositionRestrict = PaneDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="dockPosition">pane dock alignment direction</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneDockPosition dockPosition)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = dockPosition;
            DockPositionRestrict = PaneDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="dockPosition">pane dock alignment direction</param>
        /// <param name="restriction">pane dock alignment restriction</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneDockPosition dockPosition, PaneDockPositionRestrict restriction)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = dockPosition;
            DockPositionRestrict = restriction;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="dockPosition">pane dock alignment direction</param>
        /// <param name="restriction">pane dock alignment restriction</param>
        /// <param name="creation">pane creation mode</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneDockPosition dockPosition, PaneDockPositionRestrict restriction, PaneCreation creation)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = dockPosition;
            DockPositionRestrict = restriction;
            Creation = creation;
            Width = 150;
            Height = 150;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="dockPosition">pane dock alignment direction</param>
        /// <param name="restriction">pane dock alignment restriction</param>
        /// <param name="width">pane width</param>
        /// <param name="height">pane height</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneDockPosition dockPosition, PaneDockPositionRestrict restriction, int width, int height)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = dockPosition;
            DockPositionRestrict = restriction;
            Width = width;
            Height = height;
        }

        /// <summary>
        /// Creates an instance of the Attribute
        /// </summary>
        /// <param name="paneType">type of the custom task pane</param>
        /// <param name="title">pane caption</param>
        /// <param name="visible">pane visibility</param>
        /// <param name="dockPosition">pane dock alignment direction</param>
        /// <param name="restriction">pane dock alignment restriction</param>
        /// <param name="width">pane width</param>
        /// <param name="height">pane height</param>
        /// <param name="creation">pane creation mode</param>
        public CustomPaneAttribute(Type paneType, string title, bool visible, PaneDockPosition dockPosition, PaneDockPositionRestrict restriction, int width, int height, PaneCreation creation)
        {
            if (null == paneType)
                throw new ArgumentException("paneType");
            PaneType = paneType;
            Title = title;
            Visible = visible;
            DockPosition = dockPosition;
            DockPositionRestrict = restriction;
            Width = width;
            Height = height;
            Creation = creation;
        }

        #endregion
    }
}
