using System;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    public delegate void Worksheet_SelectionChangeEventHandler(NetOffice.IExcelApi.IRange target);

    [EntityType(EntityType.IsCoClass), EventSink(typeof(DocEvents_SinkHelper))]
    public interface IWorksheet : I_Worksheet, IEventBinding
    {
        event Worksheet_SelectionChangeEventHandler SelectionChangeEvent;
    }
}