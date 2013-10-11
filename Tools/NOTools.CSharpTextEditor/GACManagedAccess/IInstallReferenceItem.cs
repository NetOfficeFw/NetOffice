// taken from http://blogs.msdn.com/b/junfeng/archive/2004/09/14/229649.aspx
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("582dac66-e678-449f-aba6-6faaec8a9394")]
    internal interface IInstallReferenceItem
    {
        // A pointer to a FUSION_INSTALL_REFERENCE structure.
        // The memory is allocated by the GetReference method and is freed when
        // IInstallReferenceItem is released. Callers must not hold a reference to this
        // buffer after the IInstallReferenceItem object is released.
        // This uses the InstallReferenceOutput object to avoid allocation
        // issues with the interop layer.
        // This cannot be marshaled directly - must use IntPtr
        [PreserveSig()]
        int GetReference(
                out IntPtr pRefData,
                int flags,
                IntPtr pvReserced);
    }// IInstallReferenceItem
}
