using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.CompilerServices;
using Visio = NetOffice.VisioApi;

namespace Out_Parameters2
{
    public static class IVPageExtensions
    {
        [DefaultMember("Name"), Guid("000D0709-0000-0000-C000-000000000046"), TypeLibType(4176)]
        [ComImport]
        [ComVisible(true)]
        public interface IVPageVTable
        {
            [DispId(32)]
            [MethodImpl(MethodImplOptions.InternalCall)]
            void GetFormulas(
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I2)] [In] ref Array SID_SRCStream,
                [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] [Out] out Array formulaArray
                );
        }
            
        public static void GetFormulasEx(this Visio.IVPage page, Int16[] sID_SRCStream, out object[] formulaArray)
        {
            formulaArray = null;

            Array arg1 = sID_SRCStream as Array;
            if(null == arg1)
                arg1 = new Array[0];

            IVPageVTable proxy = page.UnderlyingObject as IVPageVTable;
            if (null != proxy)
            {
                Array formulas = null;
                proxy.GetFormulas(ref arg1, out formulas);
                if (null != formulas)
                {                    
                    formulaArray = new object[formulas.Length];
                    for (int i = 0; i < formulas.Length; i++)
                        formulaArray[i] = formulas.GetValue(i);
                }
            }
            else
                throw new InvalidCastException("Unable to cast underlying proxy into interop interface");    
        }
    }
}
