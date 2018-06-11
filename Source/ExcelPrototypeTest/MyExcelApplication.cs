using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ExcelPrototypeTest
{
    public class MyExcelApplication : NetOffice.ExcelApi.Behind.Application
    {
        public override bool Visible
        {
            get
            {
                return base.Visible;
            }
            set
            {
                if (value == false)
                    throw new ArgumentException("ooooh no sir!");
                base.Visible = value;
            }
        }

        protected override object CallArgumentValidatedPropertyGet(string name, object[] validatedArgs, ParameterModifier[] modifiers = null)
        {
            System.Console.WriteLine("Someone get the {0} property for this instance.", name);
            return base.CallArgumentValidatedPropertyGet(name, validatedArgs, modifiers);
        }
    }
}
