using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal interface IAppDomainMethod
    {
        void SetConfig(object configInstance);

        int ExecuteInDomain();
    }
}
