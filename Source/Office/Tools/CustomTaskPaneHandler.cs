using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Tools;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="paneInfo"></param>
    /// <returns></returns>
    public delegate bool CallOnCreateTaskPaneInfoHandler(TaskPaneInfo paneInfo);

    /// <summary>
    /// 
    /// </summary>
    public class CustomTaskPaneHandler
    {
        /// <summary>
        /// Analyze COMAddin custom taskpane attributes
        /// </summary>
        /// <param name="taskPanes">taskpanes you want to create</param>
        /// <param name="addinType">addin class type informations</param>
        /// <param name="callOnCreateTaskPaneInfo">callback to manipulate the process dynamicly</param>
        /// <param name="visibleStateChange">visible changed event handler</param>
        /// <param name="dockPositionStateChange">dock state changed event handler</param>
        public void ProceedCustomPaneAttributes(OfficeApi.Tools.CustomTaskPaneCollection taskPanes, Type addinType,
            CallOnCreateTaskPaneInfoHandler callOnCreateTaskPaneInfo,
            CustomTaskPane_VisibleStateChangeEventHandler visibleStateChange, 
            CustomTaskPane_DockPositionStateChangeEventHandler dockPositionStateChange)
        {
            CustomPaneAttribute[] paneAttributes = AttributeReflector.GetCustomPaneAttributes(addinType);
            foreach (CustomPaneAttribute itemPane in paneAttributes)
            {
                if (null != itemPane)
                {
                    TaskPaneInfo item = taskPanes.Add(itemPane.PaneType, itemPane.PaneType.Name);
                    if (!callOnCreateTaskPaneInfo(item))
                    {
                        item.Title = itemPane.Title;
                        item.Visible = itemPane.Visible;
                        item.DockPosition = (OfficeApi.Enums.MsoCTPDockPosition)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPosition), itemPane.DockPosition.ToString());
                        item.DockPositionRestrict = (OfficeApi.Enums.MsoCTPDockPositionRestrict)Enum.Parse(typeof(OfficeApi.Enums.MsoCTPDockPositionRestrict), itemPane.DockPositionRestrict.ToString());
                        item.Width = itemPane.Width;
                        item.Height = itemPane.Height;
                        item.Arguments = new object[] { this };
                    }

                    item.VisibleStateChange += visibleStateChange;
                    item.DockPositionStateChange += dockPositionStateChange;
                }
            }
        }


        /// <summary>
        /// Create taskpanes
        /// </summary>
        /// <typeparam name="T">Taskpane interface type from</typeparam>
        /// <typeparam name="N">Current host application</typeparam>
        /// <param name="factory">current used factory</param>
        /// <param name="ctpFactoryInst">taskpane creator given from office application</param>
        /// <param name="taskPanes">taskpane you want to create</param>
        /// <param name="taskPaneInstances">create taskpane instances</param>
        /// <param name="onError">Error callback if somthings fails</param>
        /// <param name="application">host application in base definition</param>
        public OfficeApi.ICTPFactory CreateCustomPanes<T,N>(Core factory, object ctpFactoryInst, OfficeApi.Tools.CustomTaskPaneCollection taskPanes,
            List<T> taskPaneInstances,  NetOffice.Tools.OnErrorHandler onError, COMObject application) where T: class where N:COMObject
        { 
            OfficeApi.ICTPFactory TaskPaneFactory = new NetOffice.OfficeApi.ICTPFactory(factory, null, ctpFactoryInst);
            try
            {
                foreach (TaskPaneInfo item in taskPanes)
                {
                    string title = item.Title;
                    OfficeApi.CustomTaskPane taskPane = CreateCTP(factory, TaskPaneFactory, item.Type.FullName, title, onError);
                    if (null == taskPane)
                        continue;
                    Type taskPaneType = taskPane.GetType();

                    item.Pane = taskPane;
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

                    T pane = taskPane.ContentControl as T;
                    if (null != pane)
                    {
                        taskPaneInstances.Add(pane);
                        object[] argumentArray = new object[0];

                        if (item.Arguments != null)
                            argumentArray = item.Arguments;

                        try
                        {
                            OfficeApi.Tools.ITaskPaneConnection<N> foo = pane as OfficeApi.Tools.ITaskPaneConnection<N>;
                            foo.OnConnection(application as N, taskPane, argumentArray);
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

                return TaskPaneFactory;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private OfficeApi.CustomTaskPane CreateCTP(Core factory, OfficeApi.ICTPFactory taskPaneFactory, string fullName, string title, NetOffice.Tools.OnErrorHandler onError)
        {
            OfficeApi.CustomTaskPane taskPane = null;
            try
            {
                taskPane = taskPaneFactory.CreateCTP(fullName, title) as OfficeApi.CustomTaskPane;
            }
            catch (System.Exception exception)
            {
                string message = String.Format("Unable to create {0}({1}).", fullName, title);
                System.Runtime.InteropServices.COMException wrapperException = new System.Runtime.InteropServices.COMException(message, exception);
                factory.Console.WriteException(wrapperException);
                onError(NetOffice.Tools.ErrorMethodKind.CTPFactoryAvailable, wrapperException);
            }
            return taskPane;
        }
    }
}
