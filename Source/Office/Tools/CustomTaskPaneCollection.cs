using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using NetOffice;
using NetOffice.Tools;
using NetOffice.Attributes;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using System.Diagnostics;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Wrapper class for CustomTaskPane instance, also used as creation definition if its create before CTPFactoryAvailable is called from MS-Office host application. (Best use in .ctor for creation definition)
    /// </summary>
    public class TaskPaneInfo
    {
        #region Fields

        private CustomTaskPane_VisibleStateChangeEventHandler _visibleStateChange;
        private CustomTaskPane_DockPositionStateChangeEventHandler _dockPositionStateChange;
        private Action<TaskPaneInfo> _createAction;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">Type information for the specified UserControl</param>
        /// <param name="title">title of the control</param>
        /// <param name="createAtStartup">create pane while addin startup, otherwise on demand</param>
        internal TaskPaneInfo(Type type, string title, bool createAtStartup)
        {
            ChangedProperties = new Dictionary<string, object>();
            Type = type;
            Title = title;
            CreateAtStartup = createAtStartup;
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs when task visibility is changed
        /// </summary>
        public event CustomTaskPane_VisibleStateChangeEventHandler VisibleStateChange
        {
            add
            {
                _visibleStateChange += value;
            }
            remove
            {
                _visibleStateChange -= value;
            }
        }

        /// <summary>
        /// Raise the VisibleChanged event
        /// </summary>
        internal void RaiseVisibleChanged(Office._CustomTaskPane pane)
        {
            var handler = _visibleStateChange;
            if (null != handler)
                handler(pane);
        }

        /// <summary>
        /// Occurs when dock postion state is changed
        /// </summary>
        public event CustomTaskPane_DockPositionStateChangeEventHandler DockPositionStateChange
        {
            add
            {
                _dockPositionStateChange += value;
            }
            remove
            {
                _dockPositionStateChange -= value;
            }
        }

        /// <summary>
        /// Raise the DockPositionStateChange event
        /// </summary>
        internal void RaiseDockPositionStateChanged(Office._CustomTaskPane pane)
        {
            var handler = _visibleStateChange;
            if (null != handler)
                handler(pane);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Default Height or Width if unset - 150
        /// </summary>
        public static int DefaultSize
        {
            get
            {
                return 150;
            }
        }

        /// <summary>
        /// properties was set from the client before the instance was created. The COMAddin class perfom latebind property set calls during this dictionary
        /// </summary>
		[Browsable(false), EditorBrowsable( EditorBrowsableState.Never)]
        public Dictionary<string, object> ChangedProperties { get; private set; }

        /// <summary>
        /// info about the inner taskpane instance is already created
        /// </summary>
		[Browsable(false), EditorBrowsable( EditorBrowsableState.Never)]
        public bool IsLoaded { get; set; }

        /// <summary>
        /// Determines the pane should created while addin startup, otherwise on demand
        /// </summary>
        public bool CreateAtStartup { get; private set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool Visible
        {
            get
            {
                if (IsLoaded)
                    return Pane.Visible;

                object outValue;
                if(ChangedProperties.TryGetValue("Visible", out outValue))
                    return Convert.ToBoolean(outValue);
                else
                    return false;
            }
            set
            {
                if (IsLoaded)
                    Pane.Visible = value;

                 object outValue;
                 if(ChangedProperties.TryGetValue("Visible", out outValue))
                    ChangedProperties["Visible"] = value;
                 else
                     ChangedProperties.Add("Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public int Width
        {
            get
            {
                if (IsLoaded)
                    return Pane.Width;

                object outValue;
                if(ChangedProperties.TryGetValue("Width", out outValue))
                    return Convert.ToInt32(outValue);
                else
                    return 0;
            }
            set
            {
                if (IsLoaded)
                    Pane.Width = value;

                 object outValue;
                 if(ChangedProperties.TryGetValue("Width", out outValue))
                    ChangedProperties["Width"] = value;
                 else
                     ChangedProperties.Add("Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public int Height
        {
            get
            {
                if (IsLoaded)
                    return Pane.Height;

                object outValue;
                if(ChangedProperties.TryGetValue("Height", out outValue))
                    return Convert.ToInt32(outValue);
                else
                    return 0;
            }
            set
            {
                if (IsLoaded)
                   Pane.Height = value;

                 object outValue;
                 if(ChangedProperties.TryGetValue("Height", out outValue))
                    ChangedProperties["Height"] = value;
                 else
                     ChangedProperties.Add("Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string Title
        {
            get
            {
                if (IsLoaded)
                    return Pane.Title;

                object outValue;
                if (ChangedProperties.TryGetValue("Title", out outValue))
                    return Convert.ToString(outValue);
                else
                    return "";
            }
            set
            {
                object outValue;
                if (ChangedProperties.TryGetValue("Title", out outValue))
                    ChangedProperties["Title"] = value;
                else
                    ChangedProperties.Add("Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public MsoCTPDockPosition DockPosition
        {
            get
            {
                if (IsLoaded)
                    return Pane.DockPosition;

                object outValue;
                if (ChangedProperties.TryGetValue("DockPosition", out outValue))
                    return (MsoCTPDockPosition)outValue;
                else
                    return 0;
            }
            set
            {
                if (IsLoaded)
                    Pane.DockPosition = value;

                object outValue;
                if (ChangedProperties.TryGetValue("DockPosition", out outValue))
                    ChangedProperties["DockPosition"] = value;
                else
                    ChangedProperties.Add("DockPosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public MsoCTPDockPositionRestrict DockPositionRestrict
        {
            get
            {
                if (IsLoaded)
                    return Pane.DockPositionRestrict;

                object outValue;
                if (ChangedProperties.TryGetValue("DockPositionRestrict", out outValue))
                    return (MsoCTPDockPositionRestrict)outValue;
                else
                    return 0;
            }
            set
            {
                if (IsLoaded)
                    Pane.DockPositionRestrict = value;

                object outValue;
                if (ChangedProperties.TryGetValue("DockPositionRestrict", out outValue))
                    ChangedProperties["DockPositionRestrict"] = value;
                else
                    ChangedProperties.Add("DockPositionRestrict", value);
            }
        }

        /// <summary>
        /// CustomTaskPane instance
        /// </summary>
        public NetOffice.OfficeApi.CustomTaskPane Pane { get; set; }

        /// <summary>
        /// UserControl type info
        /// </summary>
        public Type Type { get; internal set; }

		/// <summary>
        /// Additional arguments for OnConnection. The UserControl must implement ITaskPane to use it
        /// </summary>
		public object[] Arguments{ get; set; }

        /// <summary>
        /// Custom Tag as any
        /// </summary>
        public object Tag { get; set; }

        #endregion

        #region Methods / Trigger

        /// <summary>
        /// Creates the pane in office application if necessary
        /// </summary>
        /// <returns>true if pane is newly created, otherwise false</returns>
        public bool Create()
        {
            if (null == Pane && null != _createAction)
            {
                _createAction(this);
                return true;
            }
            else
            {
                return false;
            }
        }

		/// <summary>
		/// Attach the event triggers
		/// </summary>
        //[Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        internal void AssignEvents()
        {
            if (null != Pane && !Pane.IsDisposed && System.Runtime.InteropServices.Marshal.IsComObject(Pane.UnderlyingObject))
            {
                Pane.VisibleStateChangeEvent += Pane_VisibleStateChangeEvent;
                Pane.DockPositionStateChangeEvent += Pane_DockPositionStateChangeEvent;
            }
        }

        internal void SetCreateAction(Action<TaskPaneInfo> createAction)
        {
            _createAction = createAction;
        }

        private void Pane_DockPositionStateChangeEvent(_CustomTaskPane customTaskPaneInst)
        {
            try
            {
                RaiseDockPositionStateChanged(customTaskPaneInst);
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
            }
        }

        private void Pane_VisibleStateChangeEvent(_CustomTaskPane customTaskPaneInst)
        {
            try
            {
                RaiseVisibleChanged(customTaskPaneInst);
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
            }
        }

        #endregion
    }

    /// <summary>
    /// TaskCollection for COMAddin
    /// </summary>
    [DebuggerDisplay("{Count} Items")]
    public class CustomTaskPaneCollection : IEnumerable<TaskPaneInfo>
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public CustomTaskPaneCollection()
        {
            InnerList = new List<TaskPaneInfo>();
        }

        private List<TaskPaneInfo> InnerList { get; set; }

        /// <summary>
        /// Add a new child to the list
        /// </summary>
        /// <param name="taskPaneType">new child</param>
        /// <param name="title">title(caption) of the child</param>
        /// <param name="paneCreation">create at startup, otherwise on demand</param>
		/// <returns>new instance</returns>
        public TaskPaneInfo Add(Type taskPaneType, string title, PaneCreation paneCreation)
        {
			TaskPaneInfo item = new TaskPaneInfo(taskPaneType, title, paneCreation == PaneCreation.AutomaticallyAtStartup);
            InnerList.Add(item);
			return item;
        }

        /// <summary>
        /// Remove a taskpane from the collection
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        internal bool Remove(TaskPaneInfo item)
        {
            if (null != item)
                return InnerList.Remove(item);
            else
                return false;
        }

		/// <summary>
        /// Collection items count
        /// </summary>
		public int Count
		{
			get
			{
				return InnerList.Count;
			}
		}

        /// <summary>
        /// Returns an element from specified index
        /// </summary>
        /// <param name="index">specified index</param>
        /// <returns>TaskPaneInfo instance</returns>
        public TaskPaneInfo this[int index]
        {
            get
            {
                return InnerList[index];
            }
        }

        /// <summary>
        /// Returns first element with specified title or null(Nothing in Visual Basic)
        /// </summary>
        /// <param name="title">specified title</param>
        /// <returns>TaskPaneInfo instance</returns>
        public TaskPaneInfo this[string title]
        {
            get
            {
                return InnerList.FirstOrDefault(e => e.Title == title);
            }
        }

        /// <summary>
        /// Returns an Enumerator
        /// </summary>
        /// <returns>IEnumerator instance</returns>
        public IEnumerator<TaskPaneInfo> GetEnumerator()
        {
            return InnerList.GetEnumerator();
        }

        /// <summary>
        /// Returns an Enumerator
        /// </summary>
        /// <returns>IEnumerator instance</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return InnerList.GetEnumerator();
        }

        internal void TaskPaneDeleted(_CustomTaskPane pane)
        {
            var target = this.FirstOrDefault(e => e.Pane == pane);
            if(null != target)
                Remove(target);
        }
    }
}
