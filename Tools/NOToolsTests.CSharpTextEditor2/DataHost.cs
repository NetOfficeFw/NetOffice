using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NOToolsTests.CSharpTextEditor2.DataLayer;

namespace NOToolsTests.CSharpTextEditor2
{
    internal class DataHost : IDataHost
    {
        internal DataHost()
        {
            Tables = new RootListCollection();
            Local = new AccessContextCollection(Tables);
        }

        public RootListCollection Tables { get; private set; }

        public AccessContextCollection Local { get; private set; }
    }
}
