using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    public class MyExcelRange : NetOffice.ExcelApi.Behind.Range
    {
        public override object Value
        {
            get
            {
                return base.Value;
            }
            set
            {
                if (value == null)
                    value = "<Nothing>";
                base.Value = value;
            }
        }
    }
}
