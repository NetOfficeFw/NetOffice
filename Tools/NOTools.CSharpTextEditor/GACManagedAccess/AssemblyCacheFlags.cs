// taken from http://blogs.msdn.com/b/junfeng/archive/2004/09/14/229649.aspx
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    [Flags]
    internal enum AssemblyCacheFlags
    {
        GAC = 2,
    }
}
