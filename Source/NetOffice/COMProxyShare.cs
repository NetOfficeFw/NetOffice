using System;
using System.Runtime.InteropServices;

namespace NetOffice
{
    /* 
        Purpose:

        Managed proxies (System._ComObject) implement its own managed lifetime service and reference counter.
        Marshal.ReleaseComObject does NOT! decrement the remote IUnkown interface directly - 
        its decrement its own managed reference counter and 
        if the managed ref counter is == 0 then the remote IUnkown interface want be decremented.       
        (a common missunderstanding)

        If you increment the IUnkown reference counter directly(Marshal.AddRef) means 
        the RCW(System._ComObject) does not recognize that in its own managed lifetime service. 
     
        Unfortunately Microsoft spend no possibilities to influence the managed RCW lifetime service 
        for System._ComObject except of Marshal.ReleaseComObject/Marshal.FinalReleaseComObject.
        Thats why we spend this lifetime wrapper arround to have multiple 
        Netoffice wrapper instances with same RCW proxy and keep the managed proxy alive as long we need.
    */

    /// <summary>
    /// Provides shared access to managed COM proxies(System._ComObject) by implement a reference counter.  
    /// </summary>
    public class COMProxyShare
    {
        #region Nested

        /// <summary>
        /// COMProxyShare event handler after reference counter has been changed
        /// </summary>
        /// <param name="sender">Event sender</param>
        public delegate void COMProxyShareCountChangedChangedEventHandler(COMProxyShare sender);

        #endregion

        #region Fields

        /// <summary>
        /// Shared access thread lock in Aquire/Release
        /// </summary>
        private object _thisLock = new object();

        /// <summary>
        /// Reference count for _proxy
        /// </summary>
        protected volatile int _count;

        /// <summary>
        /// Com proxy as any
        /// </summary>
        protected object _proxy;

        /// <summary>
        /// Cache flag to see _proxy is disconnected
        /// </summary>
        protected bool _released;

        /// <summary>
        /// Instance is marked as enumerator provider
        /// </summary>
        private bool _isEnumerator;

        /// <summary>
        /// Used Core
        /// </summary>
        private Core _factory;

        /// <summary>
        /// Invalid proxy error message
        /// </summary>
        private static string _invalidComProxy = "Given argument isn't a com proxy.";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="factory">used core</param>
        /// <param name="proxy">com proxy as any</param>
        /// <exception cref="ArgumentNullException">throws when proxy is null</exception>
        internal COMProxyShare(Core factory, object proxy)
        {
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            if (!(proxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);
            _factory = null != factory ? factory : Core.Default;
            _isEnumerator = proxy is ICustomAdapter;
            _proxy = proxy;
            Acquire();
        }

        /// <summary>
        ///  Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="factory">used core</param>
        /// <param name="proxy">com proxy as any</param>
        /// <param name="isEnumerator">indicates proxy is an enumerator</param>
        /// <exception cref="ArgumentNullException">throws when proxy is null</exception>
        internal COMProxyShare(Core factory, object proxy, bool isEnumerator)
        {
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            if (false == isEnumerator && false == (proxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);
            _factory = null != factory ? factory : Core.Default;
            _isEnumerator = isEnumerator;
            _proxy = proxy;
            Acquire();
        }

        /// <summary>
        ///  Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="factory">used core</param>
        /// <param name="proxy">com proxy as any</param>
        /// <param name="isEnumerator">indicates proxy is an enumerator</param>
        /// <param name="suppressReleaseExceptions">ignore exceptions when release underlying managed proxy</param>
        /// <exception cref="ArgumentNullException">throws when proxy is null</exception>
        internal COMProxyShare(Core factory, object proxy, bool isEnumerator, bool suppressReleaseExceptions)
        {
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            if (false == isEnumerator && false == (proxy is MarshalByRefObject))
                throw new ArgumentException(_invalidComProxy);
            _factory = null != factory ? factory : Core.Default;
            _isEnumerator = isEnumerator;
            _proxy = proxy;
            SuppressReleaseExceptions = suppressReleaseExceptions;
            Acquire();
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs after reference counter has been changed
        /// </summary>
        public COMProxyShareCountChangedChangedEventHandler CountChanged;

        #endregion

        #region Properties

        /// <summary>
        /// Returns information the underlying proxy is already released
        /// </summary>
        public bool Released
        {
            get
            {
                return _released;
            }
        }

        /// <summary>
        /// Used Factory Core
        /// </summary>
        public Core Factory
        {
            get
            {
                return _factory;
            }
        }

        /// <summary>
        /// Underyling managed proxy(System._ComObject)
        /// </summary>
        public object Proxy
        {
            get
            {
                return _proxy;
            }
        }

        /// <summary>
        /// Ignore exceptions when release underlying managed proxy(System._ComObject)
        /// </summary>
        public virtual bool SuppressReleaseExceptions { get; set; }

        /// <summary>
        /// Instance is marked as enumerator provider
        /// </summary>
        public bool IsEnumerator
        {
            get
            {
                return _isEnumerator;
            }
        }

        /// <summary>
        /// Current Reference Count
        /// </summary>
        public int Count
        {
            get
            {
                return _count;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Increment the reference counter by 1
        /// </summary>
        public virtual void Acquire()
        {
            if (_released)
                throw new ObjectDisposedException("proxy");
            lock (_thisLock)
            {
                _count++;
            }
            RaiseCountChanged();
        }

        /// <summary>
        /// Decrement the reference counter by 1 and release the underlying proxy if counter is 0 after decrement
        /// </summary>
        /// <returns>true if underlying proxy is released, otherwise false</returns>
        public virtual bool Release()
        {
            try
            { 
                lock (_thisLock)
                {
                    _count--;
                    if (0 == _count)
                    {
                        ReleaseComObject();                    
                        _released = true;
                        return true;
                    }
                    else
                        return false;
                }                
            }
            catch(Exception exception)
            {
                Factory.Console.WriteException(exception);
                if (!SuppressReleaseExceptions)
                    throw;
                else
                    return false;
            }
            finally
            {
                RaiseCountChanged();
            }
        }

        private void ReleaseComObject()
        {
            if(_isEnumerator)
            {
                ICustomAdapter adapter = TryConvertToCustomAdapter();
                if (null != adapter)
                {
                    object adapterUnderlyingObject = AdapterGetUnderlyingObject(adapter);
                    MarshalReleaseComObject(adapterUnderlyingObject);
                }
                else
                    MarshalReleaseComObject(_proxy);
            }
            else
                MarshalReleaseComObject(_proxy);
            _proxy = null;
        }

        private static object AdapterGetUnderlyingObject(ICustomAdapter adapter)
        {
            try
            {
                return adapter.GetUnderlyingObject();
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }          
        }

        private static void MarshalReleaseComObject(object proxy)
        {
            try
            {
                Marshal.ReleaseComObject(proxy);
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        private ICustomAdapter TryConvertToCustomAdapter()
        {
            try
            {
                ICustomAdapter adapter = _proxy as ICustomAdapter;
                return adapter;
            }
            catch
            {
                // cast want throw an exception if rcw is already disconnected
                return null;
            }
        }

        private void RaiseCountChanged()
        {
            try
            {
                // CountChanged?Invoke is unsupported in previous C# versions
                if (null != CountChanged)
                    CountChanged(this);
            }
            catch
            {
                ;
            }           
        }

        #endregion

        #region Overrides
       
        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            if(_isEnumerator)
                return String.Format("COMProxyShare:{0} (Enumerator)", _count);
            else
                return String.Format("COMProxyShare:{0}", _count);
        }

        #endregion
    }
}