using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [EntityType(EntityType.IsDispatchInterface)]
    public interface I_Application : ICOMObject
    {
        bool Visible { get; set; }

        bool DisplayAlerts { get; set; }

        IWorkbooks Workbooks { get; }

        void Calculate();

        int DDEInitiate(string app, string topic);

        object _WSFunction(object arg1);

        IRange Union(IRange arg1, IRange arg2, object arg);

        void Quit();
    }
}