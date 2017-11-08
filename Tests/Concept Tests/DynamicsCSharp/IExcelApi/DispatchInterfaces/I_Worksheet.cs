using System;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [EntityType(EntityType.IsDispatchInterface)]
    public interface I_Worksheet : ICOMObject
    {
        IApplication Application { get; }

        string Name { get; set; }

        IRange Cells { get; set; }
    }
}
