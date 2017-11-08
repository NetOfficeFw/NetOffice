using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
    public interface ISheets : ICOMObject, IEnumerable<object>
    {
        IApplication Application { get; }

        Int32 Count { get; }

        [IndexProperty]
        object this[object index] { get; }

        [CustomMethod]
        object Add();

        [CustomMethod]
        object Add(object before);

        [CustomMethod]
        object Add(object before, object after);

        [CustomMethod]
        object Add(object before, object after, object count);

        [CustomMethod]
        object Add(object before, object after, object count, object type);

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);
    }
}
