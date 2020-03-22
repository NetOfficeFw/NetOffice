using System;
using System.Runtime.InteropServices;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Represents runtime availability services for a COM Proxy
    /// </summary>
    public interface ICOMObjectAvailability
    {
        /// <summary>
        /// Returns information the proxy provides a method or property.
        /// Check want be made at runtime through IDispatch interface.
        /// </summary>
        /// <param name="name">name of the entity</param>
        /// <returns>true if available, otherwise false</returns>
        /// <exception cref="AvailabilityException">Unexpected error, see inner exception(s) for details.</exception>
        bool EntityIsAvailable(string name);

        /// <summary>
        /// Returns information the proxy provides a method or property.
        /// Check want be made at runtime through IDispatch interface.
        /// </summary>
        /// <param name="name">name of the entity</param>
        /// <param name="searchType">indicate the kind of entity the caller is looking for</param>
        /// <returns>true if available, otherwise false</returns>
        /// <exception cref="AvailabilityException">Unexpected error, see inner exception(s) for details.</exception>
        bool EntityIsAvailable(string name, Availability.SupportedEntityType searchType);
    }
}