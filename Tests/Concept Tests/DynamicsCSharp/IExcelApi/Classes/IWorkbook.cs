using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [EntityType(EntityType.IsCoClass)]
    public interface IWorkbook : I_Workbook
    {
    }
}
