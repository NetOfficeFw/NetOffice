using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    class GACAssemblyComparer : IComparer<GACAssembly>
    {
        public int Compare(GACAssembly x, GACAssembly y)
        {
            return (String.Compare(x.Name, y.Name));
        }
    }
}
