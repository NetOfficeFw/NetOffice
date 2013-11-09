using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NOToolsTests.CSharpTextEditor2.DataLayer;

namespace NOToolsTests.CSharpTextEditor2
{
    public interface IDataHost
    {
        RootListCollection Tables { get; }
        AccessContextCollection Local { get; }
    }
}
