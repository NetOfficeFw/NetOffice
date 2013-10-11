// taken from http://blogs.msdn.com/b/junfeng/archive/2004/09/14/229649.aspx
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("56b1a988-7c0c-4aa2-8639-c3eb5a90226f")]
    internal interface IInstallReferenceEnum
    {
        [PreserveSig()]
        int GetNextInstallReferenceItem(
                out IInstallReferenceItem ppRefItem,
                int flags,
                IntPtr pvReserced);
    }// IInstallReferenceEnum
}
