using System;
using System.Collections.Generic;
using System.ComponentModel;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// wrapper class for CustomTaskPane instance
    /// </summary>
    public class TaskPaneInfo
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">Type information for the specified UserControl</param>
        /// <param name="title">title of the control</param>
        internal TaskPaneInfo(Type type, string title)
        {
            ChangedProperties = new Dictionary<string, object>();
            Type = type;
            Title = title;
        }

        #endregion

        #region Properties

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
        /// SupportByVersion Office 12, 14, 15
        /// Get/Set
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
        /// SupportByVersion Office 12, 14, 15
        /// Get/Set
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
        /// SupportByVersion Office 12, 14, 15
        /// Get/Set
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
        /// SupportByVersion Office 12, 14, 15
        /// Get
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
            internal set
            {
                object outValue;
                if (ChangedProperties.TryGetValue("Title", out outValue))
                    ChangedProperties["Title"] = value;
                else
                    ChangedProperties.Add("Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15
        /// Get/Set
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
        /// SupportByVersion Office 12, 14, 15
        /// Get/Set
        /// </summary>
        [SupportByVersionAttribute("Office", 12, 14, 15)]
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
        public NetOffice.OfficeApi._CustomTaskPane Pane { get; set; }

        /// <summary>
        /// UserControl type info
        /// </summary>
        public Type Type { get; internal set; }

		/// <summary>
        /// Additional Arguments for OnConnection. The UserControl must implement ITaskPane to use it
        /// </summary>
		public object[] Arguments{ get; set; }

        #endregion
    }

    /// <summary>
    /// TaskCollection for COMAddin
    /// </summary>
    public class CustomTaskPaneCollection : IEnumerable<TaskPaneInfo>
    {
        private List<TaskPaneInfo> InnerList { get; set; }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public CustomTaskPaneCollection()
        {
            InnerList = new List<TaskPaneInfo>();
        }

        /// <summary>
        /// Add a new child to the list
        /// </summary>
        /// <param name="taskPaneType">new child</param>
        /// <param name="title">title(caption) of the child</param>
        public void Add(Type taskPaneType, string title)
        {
            InnerList.Add(new TaskPaneInfo(taskPaneType, title));  
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
        /// Returns an Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<TaskPaneInfo> GetEnumerator()
        {
            return InnerList.GetEnumerator();
        }

        /// <summary>
        /// Returns an Enumerator
        /// </summary>
        /// <returns></returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}