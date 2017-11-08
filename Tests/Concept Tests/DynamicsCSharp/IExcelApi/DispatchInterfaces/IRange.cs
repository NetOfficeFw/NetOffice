using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [SyntaxBypass]
    public interface IRange_ : ICOMObject
    {
        [InvokeAs(Invoke.Property), Visible]
        string get_Address(object rowAbsolute);
        [Redirect("get_Address")]
        string Address(object rowAbsolute);

        [InvokeAs(Invoke.Property), Visible]
        string get_Address(object rowAbsolute, object columnAbsolute);
        [Redirect("get_Address")]
        string Address(object rowAbsolute, object columnAbsolute);

        [InvokeAs(Invoke.Property), Visible]
        object get_Value(object rangeValueDataType);

        [InvokeAs(Invoke.Property), Visible]
        void set_Value(object rangeValueDataType, object value);

        [Redirect("get_Value")]
        object Value(object rangeValueDataType);
    }

    [EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Property, "_Default")]
    public interface IRange : IRange_
    {
        IApplication Application { get; }

        [IndexProperty]
        IRange this[object rowIndex] { get; set; }

        [IndexProperty]
        IRange this[object rowIndex, object columnIndex] { get; set; }

        new string Address { get; }

        new object Value { get; set; }

        object Activate();
        
        [Visible, InvokeAs(Invoke.Property)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        IRange get_End(Enums.XlDirection direction);

        [Redirect("get_End")]
        IRange End(Enums.XlDirection direction);
    }
}