using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointAddin
{
    public class MyCustomTaskPane : NetOffice.OfficeApi.Behind.CustomTaskPane
    {
        public MyCustomTaskPane()
        {

        }

        public override void Dispose()
        {
            base.Dispose();
        }

        public override void Dispose(bool disposeEventBinding)
        {
            base.Dispose(disposeEventBinding);
        }
    }
}
