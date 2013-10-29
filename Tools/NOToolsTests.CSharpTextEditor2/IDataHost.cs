using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NOToolsTests.CSharpTextEditor2.DataLayer;

namespace NOToolsTests.CSharpTextEditor2
{
    public interface IDataHost
    {
        RootListDefinitionCollection Tables { get; }       
    }
}
