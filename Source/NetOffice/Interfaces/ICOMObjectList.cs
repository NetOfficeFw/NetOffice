using System.Collections.Generic;

namespace NetOffice.Interfaces
{
    /// <summary>
    /// Represents the list that holds the COM objects.
    /// </summary>
    /// <seealso cref="IList{ICOMObject}" />
    public interface ICOMObjectList : IList<ICOMObject>
    {
    }
}