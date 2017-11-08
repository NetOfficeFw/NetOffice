using System;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
    public interface IWorkbooks : ICOMObject, IEnumerable<IWorkbook>
    {
        [IndexProperty]
        IWorkbook this[object index] { get; }

        IWorkbook Add(object template);

        IWorkbook Add();

        void Close();
    }
}
