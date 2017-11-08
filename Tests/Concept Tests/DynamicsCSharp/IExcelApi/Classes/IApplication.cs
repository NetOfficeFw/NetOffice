using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    public delegate void IApplication_NewWorkbookEventHandler(NetOffice.IExcelApi.IWorkbook wb);
  
    [ComProgId("Excel.Application")]
    [EntityType(EntityType.IsCoClass), EventSink(typeof(AppEvents_SinkHelper))]
    public interface IApplication : I_Application, IEventBinding
    {
        event IApplication_NewWorkbookEventHandler NewWorkbookEvent;
    }
}