using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// Provides remote comparsion 2 proxies points to the same instance on the server
    /// </summary>
    public static class RemoteComparsion
    {
        /// <summary>
        /// the well know IUnknown Interface ID
        /// </summary>
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// Determine 2 proxies represents the same object on COM remote server
        /// </summary>
        /// <param name="obj1">object 1 to compare</param>
        /// <param name="obj2">object 2 to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool EqualsOnServer(object obj1, object obj2)
        {
            return EqualsOnServer(obj1 as ICOMObject, obj2 as ICOMObject);
        }

        /// <summary>
        /// Determine 2 proxies represents the same object on COM remote server
        /// </summary>
        /// <param name="obj1">object 1 to compare</param>
        /// <param name="obj2">object 2 to compare</param>
        /// <returns>true if equal, otherwise false</returns>
        public static bool EqualsOnServer(ICOMObject obj1, ICOMObject obj2)
        {
            if (obj1.IsCurrentlyDisposing || obj1.IsDisposed)
                return ReferenceEquals(obj1, obj2);

            if (Object.ReferenceEquals(obj2, null))
                return false;

            IntPtr outValueA = IntPtr.Zero;
            IntPtr outValueB = IntPtr.Zero;
            IntPtr ptrA = IntPtr.Zero;
            IntPtr ptrB = IntPtr.Zero;
            try
            {
                ptrA = Marshal.GetIUnknownForObject(obj1.UnderlyingObject);
                int hResultA = Marshal.QueryInterface(ptrA, ref IID_IUnknown, out outValueA);

                ptrB = Marshal.GetIUnknownForObject(obj2.UnderlyingObject);
                int hResultB = Marshal.QueryInterface(ptrB, ref IID_IUnknown, out outValueB);

                return (hResultA == 0 && hResultB == 0 && ptrA == ptrB);
            }
            catch (Exception exception)
            {
                obj1.Console.WriteException(exception);
                throw exception;
            }
            finally
            {
                if (IntPtr.Zero != ptrA)
                    Marshal.Release(ptrA);

                if (IntPtr.Zero != outValueA)
                    Marshal.Release(outValueA);

                if (IntPtr.Zero != ptrB)
                    Marshal.Release(ptrB);

                if (IntPtr.Zero != outValueB)
                    Marshal.Release(outValueB);
            }
        }
    }
}
