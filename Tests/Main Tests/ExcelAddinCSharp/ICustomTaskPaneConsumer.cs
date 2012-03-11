using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelAddinCSharp
{
    [Guid("000C033E-0000-0000-C000-000000000046"), ComVisible(true)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.IDispatch)] object CTPFactoryInst);
    }
}
