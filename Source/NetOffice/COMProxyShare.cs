using System;
using System.Runtime.InteropServices;

namespace NetOffice
{
    /* 
        Purpose:

        Managed proxies (System._ComObject) implement its own managed lifetime service and reference counter.
        Marshal.ReleaseComObject does NOT! decrement the remote IUnkown interface directly - 
        its decrement its own managed reference counter and 
        if the managed ref counter is <= 0 then the remote IUnkown interface want be decremented.       
        (a common missunderstanding)

        If you increment the IUnkown reference counter directly(Marshal.AddRef) means 
        the RCW(System._ComObject) does not recognize that in its own managed lifetime service. 
        (Moreover there is no Marshal.RemoveRef - a broken implementation for a passionate COM developer)

        (You can try Marshal.GetIUnknownForObject&Marshal.AddRef and then call Marshal.ReleaseComObject 2x times and see what happen)

        Unfortunately Microsoft spend no possibilities to influence the managed RCW lifetime service 
        for System._ComObject except of Marshal.ReleaseComObject/Marshal.FinalReleaseComObject.
        Thats why we spend this lifetime wrapper arround to have multiple 
        Netoffice wrapper instances with same RCW proxy and keep the managed proxy alive as long we need.
    */

    /// <summary>
    /// Provides shared access to managed COM proxies(System._ComObject) by implement a reference counter.
    /// COMProxyShare does not provide any thread safe operations because 
    /// all essential NetOffice.Core operations are thread-safe and its not intended to use COMProxyShare outside from Netoffice infrastructure.
    /// </summary>
    public class COMProxyShare
    {
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
        /// Instance is an enumerator provider
        /// </summary>
        private bool _isEnumerator;

        /// <summary>
        /// Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="proxy">com proxy as any</param>
        /// <exception cref="ArgumentNullException">throws when proxy is null</exception>
        internal COMProxyShare(object proxy)
        {    
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            _isEnumerator = proxy is ICustomAdapter;
            _proxy = proxy;
            Aquire();
        }

        /// <summary>
        ///  Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="proxy">com proxy as any</param>
        /// <param name="isEnumerator">indicates proxy is an enumerator</param>
        /// <exception cref="ArgumentNullException">throws when proxy is null</exception>
        internal COMProxyShare(object proxy, bool isEnumerator)
        {
            if (null == proxy)
                throw new ArgumentNullException("proxy");
            _isEnumerator = isEnumerator;
            _proxy = proxy;
            Aquire();
        }

        /// <summary>
        /// Returns information the underlying proxy is already released
        /// </summary>
        public virtual bool Released
        {
            get
            {
                return _released;
            }
        }

        /// <summary>
        /// Underyling managed proxy(System._ComObject)
        /// </summary>
        public virtual object Proxy
        {
            get
            {
                return _proxy;
            }
        }

        /// <summary>
        /// Increment the reference counter by 1
        /// </summary>
        public virtual void Aquire()
        {
            if (_released)
                throw new ObjectDisposedException("proxy");
            _count++;
        }

        /// <summary>
        /// Decrement the reference counter by 1 and release the proxy if counter is 0 after decrement
        /// </summary>
        /// <returns>true if underlying proxy is released, otherwise false</returns>
        public virtual bool Release()
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
    }
}