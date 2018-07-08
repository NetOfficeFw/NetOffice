using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections;
using NetOffice.Tools.Isolation;

namespace NetOffice.Tools
{
    /// <summary>
    /// Encapsulate generic addin services
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddinBase : ICustomQueryInterface, IManagedInnerAddin
    {
        #region Fields

        /// <summary>
        /// Set in ctor first to measure the time from creation to OnStartupComplete
        /// </summary>
        protected DateTime _creationTime;

        /// <summary>
        /// Type cache field
        /// </summary>
        private Type _type;

        /// <summary>
        /// Static visual styles lock
        /// </summary>
        private object _lock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public COMAddinBase()
        {
            _creationTime = DateTime.Now;
            EnableVisualStyles();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Host Application Instance
        /// </summary>
        public abstract ICOMObject AppInstance { get; }

        /// <summary>
        /// Current asscociated Core
        /// </summary>
        public abstract Core Factory { get; }

        /// <summary>
        /// Elapsed time in milliseconds from instance creation until OnStartupComplete event
        /// </summary>
        public TimeSpan LoadingTimeElapsed { get; protected set; }

        /// <summary>
        /// Type Information of the instance
        /// </summary>
        public Type Type
        {
            get
            {
                if (null == _type)
                    _type = GetType();
                return _type;
            }
        }

        /// <summary>
        /// Instance managed root com objects
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public IEnumerable Roots { get; protected set; }

        /// <summary>
        /// Outer COM Shim if addin is isolated by a Shim
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        protected internal IShimHost ShimHost { get; private set; }

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

        #region IManagedInnerAddin

        /// <summary>
        /// Set an unmanaged aggregator to a managed addin instance
        /// </summary>
        /// <param name="aggregator">outer aggregator</param>
        void IManagedInnerAddin.SetParent(IShimHost aggregator)
        {
            ShimHost = aggregator;
        }

        void IManagedInnerAddin.ReloadNotification(string custom)
        {
            try
            {
                RecieveCustomData(custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        /// <summary>
        /// Recieve custom data from an addin update handler
        /// </summary>
        /// <param name="custom">custom data as any</param>
        protected virtual void RecieveCustomData(string custom)
        {

        }

        #endregion

        #region Methods

        /// <summary>
        /// Cleanup instance resources
        /// </summary>
        protected internal virtual void CleanUp()
        {
            if(null != ShimHost)
            {
                Marshal.ReleaseComObject(ShimHost);
                ShimHost = null;
            }
        }

        /// <summary>
        /// Call System.Windows.Forms.Application.EnableVisualStyles
        /// </summary>
        protected internal virtual void EnableVisualStyles()
        {
            lock (_lock)
            {
                if (System.Windows.Forms.Application.VisualStyleState == System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled)
                    System.Windows.Forms.Application.EnableVisualStyles();
            }
        }

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
                if(null != instance && null != type)
                    result = Marshal.GetComInterfaceForObject(instance, type, CustomQueryInterfaceMode.Ignore);
            }
            catch (Exception)
            {
                ;
            }
            return result;
        }

        #endregion
    }
}
