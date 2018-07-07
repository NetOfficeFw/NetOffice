using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Tools;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Tools;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// TaskPane OnCreate EventHandler
    /// </summary>
    /// <param name="paneInfo">new created pane</param>
    public delegate bool CallOnCreateTaskPaneInfoHandler(TaskPaneInfo paneInfo);

    /// <summary>
    /// Manage Taskpane creation process
    /// </summary>
    public class CustomTaskPaneHandler
    {
        /// <summary>
        /// Analyze COMAddin custom taskpane attributes
        /// </summary>
        /// <param name="factory">current used factory</param>
        /// <param name="taskPaneFactory">taskpane creator given from office application</param>
        /// <param name="taskPanes">taskpanes you want to create</param>
        /// <param name="onError">Error callback if somthings fails</param>
        /// <param name="application">host application in base definition</param>
        /// <param name="addinType">addin class type informations</param>
        /// <param name="addin">addin instance</param>
        /// <param name="callOnCreateTaskPaneInfo">callback to manipulate the process dynamicly</param>
        /// <param name="visibleStateChange">visible changed event handler</param>
        /// <param name="dockPositionStateChange">dock state changed event handler</param>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public void ProceedCustomPaneAttributes(Core factory, OfficeApi.ICTPFactory taskPaneFactory, OfficeApi.Tools.CustomTaskPaneCollection taskPanes,
             NetOffice.Tools.OnErrorHandler onError, ICOMObject application,
            Type addinType, IOfficeCOMAddin addin,
            CallOnCreateTaskPaneInfoHandler callOnCreateTaskPaneInfo,
            CustomTaskPane_VisibleStateChangeEventHandler visibleStateChange,
            CustomTaskPane_DockPositionStateChangeEventHandler dockPositionStateChange)
        {
            try
            {
                var paneAttributes = NetOffice.Attributes.AttributeExtensions.GetCustomAttributes<CustomPaneAttribute>(addinType);
                foreach (CustomPaneAttribute itemPane in paneAttributes)
                {
                    if (null != itemPane)
                    {
                        TaskPaneInfo item = taskPanes.Add(itemPane.PaneType, itemPane.PaneType.Name, itemPane.Creation);
                        if (!item.CreateAtStartup)
                        {
                            Action<TaskPaneInfo> method = delegate (TaskPaneInfo info)
                            {
                                CreateCustomPane(info, factory, taskPaneFactory, taskPanes, onError, application);
                            };
                            item.SetCreateAction(method);
                        }

                        item.Title = itemPane.Title;
                        item.Visible = itemPane.Visible;
                        item.DockPosition = (OfficeApi.Enums.MsoCTPDockPosition)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPosition), itemPane.DockPosition.ToString());
                        item.DockPositionRestrict = (OfficeApi.Enums.MsoCTPDockPositionRestrict)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPositionRestrict), itemPane.DockPositionRestrict.ToString());
                        item.Width = itemPane.Width;
                        item.Height = itemPane.Height;
                        item.Arguments = new object[] { addin, this };
                        if (callOnCreateTaskPaneInfo(item))
                        {
                            item.VisibleStateChange += visibleStateChange;
                            item.DockPositionStateChange += dockPositionStateChange;
                        }
                        else
                        {
                            taskPanes.Remove(item);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                if (null != onError)
                    onError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        /// <summary>
        /// Create taskpanes
        /// </summary>
        /// <param name="factory">current used factory</param>
        /// <param name="taskPaneFactory">taskpane creator given from office application</param>
        /// <param name="taskPanes">taskpane you want to create</param>
        /// <param name="onError">Error callback if somthings fails</param>
        /// <param name="application">host application in base definition</param>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        public void CreateCustomPanes(Core factory, OfficeApi.ICTPFactory taskPaneFactory, OfficeApi.Tools.CustomTaskPaneCollection taskPanes,
            NetOffice.Tools.OnErrorHandler onError, ICOMObject application)
        {
            if (null == factory)
                return;
            if (null == taskPaneFactory)
                return;
            if (null == taskPanes)
                return;
            if (null == application)
                return;

            try
            {
                foreach (TaskPaneInfo item in taskPanes)
                {
                    if (item.CreateAtStartup)
                    {
                        CreateCustomPane(item, factory, taskPaneFactory, taskPanes, onError, application);
                    }
                }
            }
            catch (Exception exception)
            {
                if(null != onError)
                    onError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        private void CreateCustomPane(TaskPaneInfo item, Core factory, OfficeApi.ICTPFactory taskPaneFactory, OfficeApi.Tools.CustomTaskPaneCollection taskPanes,
           NetOffice.Tools.OnErrorHandler onError, ICOMObject application)
        {
            try
            {
                if (null == factory)
                    return;
                if (null == taskPaneFactory)
                    return;
                if (null == taskPanes)
                    return;
                if (null == application)
                    return;
                if (null == item)
                    return;
                if (null != item.Pane)
                    return;

                string title = item.Title;
                OfficeApi.CustomTaskPane taskPane = CreateCTP(taskPaneFactory, item.Type.FullName, title, onError);
                if (null == taskPane)
                {
                    return;
                }

                Type taskPaneType = taskPane.GetType();
                item.Pane = taskPane;
                taskPane.AfterDelete += taskPanes.TaskPaneDeleted;

                item.AssignEvents();
                item.IsLoaded = true;

                switch (taskPane.DockPosition)
                {
                    case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionLeft:
                    case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionRight:
                        taskPane.Width = item.Width >= 0 ? item.Width : TaskPaneInfo.DefaultSize;
                        break;
                    case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionTop:
                    case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionBottom:
                        taskPane.Height = item.Height >= 0 ? item.Height : TaskPaneInfo.DefaultSize;
                        break;
                    case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionFloating:
                        item.Width = item.Width >= 0 ? item.Width : TaskPaneInfo.DefaultSize;
                        taskPane.Height = item.Height >= 0 ? item.Height : TaskPaneInfo.DefaultSize;
                        break;
                    default:
                        break;
                }

                OfficeApi.Tools.ITaskPane pane = taskPane.ContentControl as OfficeApi.Tools.ITaskPane;
                if (null != pane)
                {
                    object[] argumentArray = new object[0];

                    if (item.Arguments != null)
                        argumentArray = item.Arguments;

                    try
                    {
                        OfficeApi.Tools.ITaskPane foo = pane as OfficeApi.Tools.ITaskPane;
                        if (null != foo)
                            foo.OnConnection(application, taskPane, argumentArray);
                    }
                    catch (Exception exception)
                    {
                        factory.Console.WriteException(exception);
                    }
                }

                foreach (KeyValuePair<string, object> property in item.ChangedProperties)
                {
                    if (property.Key == "Title")
                        continue;

                    try
                    {
                        if (property.Key == "Width") // avoid to set width in top and bottom align
                        {
                            object outValue = null;
                            item.ChangedProperties.TryGetValue("DockPosition", out outValue);
                            if (null != outValue)
                            {

                                OfficeApi.Enums.MsoCTPDockPosition position = (OfficeApi.Enums.MsoCTPDockPosition)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPosition), outValue.ToString());
                                if (position == OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionTop || position == OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionBottom)
                                    continue;
                            }
                        }

                        if (property.Key == "Height")   // avoid to set height in left and right align
                        {
                            object outValue = null;
                            item.ChangedProperties.TryGetValue("DockPosition", out outValue);
                            if (null == outValue)
                                outValue = OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionRight; // NetOffice default position if unset

                            OfficeApi.Enums.MsoCTPDockPosition position = (OfficeApi.Enums.MsoCTPDockPosition)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPosition), outValue.ToString());
                            if (position == OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionLeft || position == OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionRight)
                                continue;
                        }

                        taskPaneType.InvokeMember(property.Key, System.Reflection.BindingFlags.SetProperty, null, taskPane, new object[] { property.Value });
                    }
                    catch
                    {
                        factory.Console.WriteLine("Failed to set TaskPane property {0}", property.Key);
                    }
                }
            }
            catch (Exception exception)
            {
                if (null != onError)
                    onError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        private OfficeApi.CustomTaskPane CreateCTP(OfficeApi.ICTPFactory taskPaneFactory, string fullName, string title, NetOffice.Tools.OnErrorHandler onError)
        {
            OfficeApi.CustomTaskPane taskPane = null;
            try
            {
                taskPane = taskPaneFactory.CreateCTP(fullName, title) as OfficeApi.CustomTaskPane;
            }
            catch (System.Exception exception)
            {
                if (null != onError)
                {
                    string message = String.Format("Unable to create {0}(Title:{1}).", fullName, title);
                    System.Runtime.InteropServices.COMException wrapperException = new NetOffice.Exceptions.NetOfficeCOMException(message, exception);
                    taskPaneFactory.Factory.Console.WriteException(wrapperException);
                    onError(NetOffice.Tools.ErrorMethodKind.CTPFactoryAvailable, wrapperException);
                }
            }
            return taskPane;
        }
    }
}
