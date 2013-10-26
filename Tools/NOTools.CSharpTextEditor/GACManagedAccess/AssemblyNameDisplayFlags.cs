// taken from http://blogs.msdn.com/b/junfeng/archive/2004/09/14/229649.aspx
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    [Flags]
    internal enum AssemblyNameDisplayFlags
    {
        VERSION = 0x01,
        CULTURE = 0x02,
        PUBLIC_KEY_TOKEN = 0x04,
        PROCESSORARCHITECTURE = 0x20,
        RETARGETABLE = 0x80,
        // This enum will change in the future to include
        // more attributes.
        ALL = VERSION
                                    | CULTURE
                                    | PUBLIC_KEY_TOKEN
                                    | PROCESSORARCHITECTURE
                                    | RETARGETABLE
    }
}
