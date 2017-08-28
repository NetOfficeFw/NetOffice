using System;
using System.Runtime.InteropServices;

namespace NetOffice
{
    /// <summary>
    /// Provides simple shared access to RCW COM proxies by implement a reference counter.
    /// COMProxyShare do not provide any thread safe operations.
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
        /// Creates an instance of the class an aquire the given proxy
        /// </summary>
        /// <param name="proxy">com proxy as any</param>
        internal COMProxyShare(object proxy)
        {
            _isEnumerator = proxy is ICustomAdapter;
            _proxy = proxy;
            Aquire();
        }

        /// <summary>
        ///  Creates an instance of the class and aquire the given proxy
        /// </summary>
        /// <param name="proxy">com proxy as any</param>
        /// <param name="isEnumerator">indicates proxy is an enumerator</param>
        internal COMProxyShare(object proxy, bool isEnumerator)
        {
            _isEnumerator = isEnumerator;
            _proxy = proxy;
            Aquire();
        }

        private bool _isEnumerator;

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
        /// Underyling RCW proxy
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
        /// Decrement the reference counter by 1 and release the proxy if 0
        /// </summary>
        /// <returns>true if underlying rcw is disconnected, otherwise false</returns>
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
                if(null != adapter)
                    Marshal.ReleaseComObject(adapter.GetUnderlyingObject());
            }
            else
                MarshalReleaseComObject(_proxy);
            _proxy = null;
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