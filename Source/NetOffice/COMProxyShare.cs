using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Handle the shared access RCW. Not intended to use from client callers.
    /// </summary>
    public interface ICOMProxyShareProvider
    {
        /// <summary>
        /// Returns the inner proxy shared access handler
        /// </summary>
        /// <returns>shared proxy</returns>
        COMProxyShare GetProxyShare();

        /// <summary>
        /// Set the inner proxy shared access handler.
        /// The method want aquire the share 1x times
        /// </summary>
        /// <param name="share">target share</param>
        void SetProxyShare(COMProxyShare share);
    }

    /// <summary>
    /// Provides simple shared access to RCW COM proxies
    /// </summary>
    public class COMProxyShare
    {   
        private int _count;
        private object _proxy;
        private bool _released;

        internal COMProxyShare(object proxy)
        {
            _proxy = proxy;
            Aquire();
        }

        internal COMProxyShare(object proxy, bool isEnumerator)
        {
            // isEnumerator is ignored, see ReleaseComObject
            _proxy = proxy;
            Aquire();
        }

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
        /// Underyling RCW proxy
        /// </summary>
        public object Proxy
        {
            get
            {
                return _proxy;
            }
        }

        /// <summary>
        /// Increment the reference counter by 1
        /// </summary>
        public void Aquire()
        {
            if (_released)
                throw new ObjectDisposedException("proxy");
            _count++;
        }

        /// <summary>
        /// Decrement the reference counter by 1 and release the proxy if 0
        /// </summary>
        /// <returns></returns>
        public bool Release()
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
            // we ignore _isEnumerator here so far and do try convert
            // want to change if its cause issues or performance problems
            ICustomAdapter adapter = TryConvertToCustomAdapter();
            if (null != adapter)
            {
                Marshal.ReleaseComObject(adapter.GetUnderlyingObject());
                Marshal.ReleaseComObject(_proxy);
            }
            else
                Marshal.ReleaseComObject(_proxy);
            _proxy = null;
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
                // cast want throw an exception if RCW is already released
                return null;
            }
        }
    }
}