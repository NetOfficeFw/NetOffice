using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.ComponentModel;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Native;
using NetOffice.OfficeApi.Tools;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// CustomTaskPaneConsumer base to seperate pane logics from addin connect
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class CustomTaskPaneConsumer : NetOffice.OfficeApi.Native.ICustomTaskPaneConsumer, ICustomQueryInterface
    {
        #region Fields

        private Type _type;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">addin owner instance</param>
        /// <exception cref="ArgumentNullException">parent is null(Nothing in Visual Basic)</exception>
        public CustomTaskPaneConsumer(IOfficeCOMAddin parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Addin Owner Instance
        /// </summary>
        public IOfficeCOMAddin Parent { get; set; }

        /// <summary>
        /// TaskPaneFactory to create custom task panes
        /// </summary>
        public Office.ICTPFactory TaskPaneFactory { get; set; }

        /// <summary>
        /// ITaskPane Instances
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        protected IEnumerable<ITaskPane> TaskPaneInstances
        {
            get
            {
                List<ITaskPane> result = new List<ITaskPane>();
                foreach (var item in Parent.TaskPanes)
                {
                    ITaskPane match = item.Pane as ITaskPane;
                    if (null != match)
                        result.Add(match);
                }
                return result.ToArray();
            }
        }

        /// <summary>
        /// Used Factory Core
        /// </summary>
        protected internal Core Factory
        {
            get
            {
                return null != Parent ? Parent.Factory : Core.Default;
            }
        }

        /// <summary>
        /// Instance Type (cached)
        /// </summary>
        protected internal Type Type
        {
            get
            {
                if (null == _type)
                    _type = GetType();
                return _type;
            }
        }

        #endregion

        #region ICustomTaskPaneConsumer

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="CTPFactoryInst">factory proxy from host application</param>
        void ICustomTaskPaneConsumer.CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
                if (null == CTPFactoryInst)
                {
                    Factory.Console.WriteLine("Warning: null argument recieved in CTPFactoryAvailable. argument name: CTPFactoryInst");
                    return;
                }

                OfficeApi.ICTPFactory taskPaneFactory = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.ICTPFactory>(null, CTPFactoryInst, typeof(NetOffice.OfficeApi.ICTPFactory));
                OnCTPFactoryAvailable(taskPaneFactory);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="ctpFactoryInst">factory proxy from host application</param>
        protected internal virtual void OnCTPFactoryAvailable(OfficeApi.ICTPFactory ctpFactoryInst)
        {
            CustomTaskPaneHandler paneHandler = new CustomTaskPaneHandler();
            paneHandler.ProceedCustomPaneAttributes(Factory, ctpFactoryInst, Parent.TaskPanes, OnError, Parent.Application, Type, Parent, CallOnCreateTaskPaneInfo, AttributePane_VisibleStateChange, AttributePane_DockPositionStateChange);
            paneHandler.CreateCustomPanes(Factory, ctpFactoryInst, Parent.TaskPanes, OnError, Parent.Application);
        }

        #endregion

        #region ICustomQueryInterface

        /// <summary>
        /// Returns an interface according to a specified interface ID
        /// </summary>
        /// <param name="iid">the GUID of the requested interface</param>
        /// <param name="ppv">a reference to the requested interface, when this method returns</param>
        /// <returns>one of the enumeration values that indicates whether a custom implementation of IUnknown::QueryInterface was used</returns>
        CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            CustomQueryInterfaceResult result = CustomQueryInterfaceResult.NotHandled;
            Type type = null;
            object instance = null;

            if (QueryInterface(iid, ref type, ref instance) ||
                QueryDefaultInterface(iid, ref type, ref instance))
            {
                ppv = TryGetComInterfaceForObject(instance, type);
                result = CustomQueryInterfaceResult.Handled;
            }

            return result;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Overrides QueryInterface default behavior
        /// </summary>
        /// <param name="interfaceId">target interface id</param>
        /// <param name="type">out argument - interface type</param>
        /// <param name="instance">out argument - instance that implements target interface</param>
        /// <returns>true if handle, otherwise false</returns>
        /// <remarks>this method allows to seperate interfaces from addin connect class</remarks>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        protected internal virtual bool QueryInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            return false;
        }

        private bool QueryDefaultInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            // currently not implemented
            return false;
        }

        private IntPtr TryGetComInterfaceForObject(object instance, Type type)
        {
            IntPtr result = IntPtr.Zero;
            try
            {
                if (null != instance && null != type)
                    result = Marshal.GetComInterfaceForObject(instance, type, CustomQueryInterfaceMode.Ignore);
            }
            catch (Exception)
            {
                ;
            }
            return result;
        }

        #endregion

        #region Trigger

        /// <summary>
        /// Called after any visibility changes
        /// </summary>
        /// <param name="customTaskPaneInst">pane instance</param>
        protected internal virtual void TaskPaneVisibleStateChanged(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {

        }

        /// <summary>
        /// Called after any position changes but not for size changes
        /// </summary>
        /// <param name="customTaskPaneInst">pane instance</param>
        protected internal virtual void TaskPaneDockStateChanged(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {

        }

        /// <summary>
        /// The method is called while the CustomPane attribute is processed
        /// </summary>
        /// <param name="paneInfo">pane definition</param>
        /// <returns>true if pane should be create, otherwise false</returns>
        protected internal virtual bool OnCreateTaskPaneInfo(TaskPaneInfo paneInfo)
        {
            return true;
        }

        /// <summary>
        /// Custom error handler
        /// </summary>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnError(ErrorMethodKind methodKind, Exception exception)
        {

        }

        private void CallTaskPaneVisibleStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {
            try
            {
                foreach (TaskPaneInfo item in Parent.TaskPanes)
                {
                    if (item.Pane.UnderlyingObject == customTaskPaneInst.UnderlyingObject)
                    {
                        try
                        {
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
                            if (null != target)
                            {
                                try
                                {
                                    target.OnVisibleStateChanged(item.Pane.Visible);
                                }
                                catch (Exception exception)
                                {
                                    Factory.Console.WriteException(exception);
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            Factory.Console.WriteException(exception);
                        }
                    }
                }
                TaskPaneVisibleStateChanged(customTaskPaneInst);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        private void CallTaskPaneDockPositionStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {
            try
            {
                foreach (TaskPaneInfo item in Parent.TaskPanes)
                {
                    if (item.Pane.UnderlyingObject == customTaskPaneInst.UnderlyingObject)
                    {
                        try
                        {
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
                            if (null != target)
                            {
                                try
                                {
                                    target.OnDockPositionChanged(item.Pane.DockPosition);
                                }
                                catch (Exception exception)
                                {
                                    Factory.Console.WriteException(exception);
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            Factory.Console.WriteException(exception);
                        }
                    }
                }
                TaskPaneDockStateChanged(customTaskPaneInst);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        private void AttributePane_VisibleStateChange(NetOffice.OfficeApi._CustomTaskPane CustomTaskPaneInst)
        {
            try
            {
                CallTaskPaneVisibleStateChange(CustomTaskPaneInst);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        private void AttributePane_DockPositionStateChange(Office._CustomTaskPane CustomTaskPaneInst)
        {
            try
            {
                CallTaskPaneDockPositionStateChange(CustomTaskPaneInst);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        private bool CallOnCreateTaskPaneInfo(TaskPaneInfo paneInfo)
        {
            try
            {
                return OnCreateTaskPaneInfo(paneInfo);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CTPFactoryAvailable, exception);
                return false;
            }
        }

        #endregion
    }
}
