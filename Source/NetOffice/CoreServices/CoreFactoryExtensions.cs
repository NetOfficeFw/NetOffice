using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;

using NetOffice.Exceptions;

namespace NetOffice.CoreServices
{
    internal static class CoreFactoryExtensions
    {
        /// <summary>
        /// See Attributes\DuplicateAttribute.cs for further informations
        /// </summary>
        private static Guid[] _duplicateTypes = new Guid[]
            {
                new Guid("000C0310-0000-0000-C000-000000000046"),
                new Guid("000C0311-0000-0000-C000-000000000046"),
                new Guid("000C0312-0000-0000-C000-000000000046"),
                new Guid("000C0314-0000-0000-C000-000000000046"),
                new Guid("000C0317-0000-0000-C000-000000000046"),
                new Guid("000C0318-0000-0000-C000-000000000046"),
                new Guid("000C0319-0000-0000-C000-000000000046"),
                new Guid("000C031A-0000-0000-C000-000000000046"),
                new Guid("000C031B-0000-0000-C000-000000000046"),
                new Guid("000C031F-0000-0000-C000-000000000046"),
                new Guid("000C0321-0000-0000-C000-000000000046"),
                new Guid("000C036E-0000-0000-C000-000000000046"),
                new Guid("000C036F-0000-0000-C000-000000000046"),
                new Guid("000C0370-0000-0000-C000-000000000046"),
                new Guid("000C0398-0000-0000-C000-000000000046")
            };

        /// <summary>
        /// Get wrapper class factory info
        /// </summary>
        /// <param name="value">core to use</param>
        /// <param name="caller">calling instance</param>
        /// <param name="comProxy">new created proxy</param>
        /// <param name="componentId">com proxy component id</param>
        /// <param name="typeId">com proxy type id</param>
        /// <param name="throwException">throw exception if no info found or return null</param>
        /// <returns>factory info from corresponding assembly</returns>
        internal static ITypeFactory GetTypeFactory(this Core value, ICOMObject caller, 
            object comProxy, Guid componentId, Guid typeId, bool throwException)
        {
            if (value.InternalFactories.FactoryAssemblies.Count == 0)
                return null;

            if (null != caller && typeId.IsDuplicateType())
            {
                // special case: if its a known duplicated type
                // we prefer to use the type from the caller component
                foreach (ITypeFactory item in value.InternalFactories.FactoryAssemblies)
                {
                    if (item.FactoryName != caller.InstanceComponentName)
                        continue;
                    if (componentId.Equals(item.ComponentID))
                        return item;
                }
            }
            else
            {
                foreach (ITypeFactory item in value.InternalFactories.FactoryAssemblies)
                {
                    if (componentId.Equals(item.ComponentID))
                        return item;
                }
            }

            string className = ComTypes.TypeDescriptor.GetClassName(comProxy);

            // failback because some types was multiple defined by its type id (not allowed in COM but in fact MS do this)
            // list of known multiple defined types is available on Attributes\DuplicateAttribute.cs
            foreach (ITypeFactory item in value.InternalFactories.FactoryAssemblies)
            {
                bool hasComponentID = null != item.ComponentID ? item.ComponentID.Equals(componentId) : false;
                if (item.ContainsType(className) && hasComponentID)
                {
                    value.Console.WriteLine("Failback factory {0}=>{1}recieved.", item.Assembly.FullName, className);
                    return item;
                }
            }

            if (throwException)
            {
                string message = string.Format("Class {0}:{1} not found in loaded NetOffice Assemblies{2}", componentId, className, Environment.NewLine);
                message += string.Format("Currently loaded NetOfficeApi Assemblies{0}", Environment.NewLine);
                foreach (ITypeFactory item in value.InternalFactories.FactoryAssemblies)
                    message += string.Format("Loaded NetOffice Assembly:{0} {1}{2}", item.ComponentID, item.Assembly.FullName, Environment.NewLine);

                throw new FactoryException(message);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns information the type is a known external duplicate
        /// </summary>
        /// <param name="value">type id</param>
        /// <returns>true if duplicate, otherwise false</returns>
        internal static bool IsDuplicateType(this Guid value)
        {
            return _duplicateTypes.Contains(value);
        }

        /// <summary>
        /// Performs GetTypeInfo on IDispatch
        /// Handle the strange cast behavior - see remarks.
        /// </summary>
        /// <remarks>
        /// Seems to be cast to IDispatch never failed
        /// even the instance behind comProxy doesnt implement the interface.
        /// If its failed to cast, an InvalidCastException occurs
        /// while first use the interface.
        /// The method catch arround here and throws IDispatchNotImplementedException
        /// to signalize the missing IDispatch suport.
        /// </remarks>
        /// <param name="dispatcher">given IDispatch as any </param>
        /// <returns>type informations or null if dispatcher argument is null</returns>
        internal static COMTypes.ITypeInfo GetTypeInfo(this IDispatch dispatcher)
        {
            if (null == dispatcher)
                return null;
            try
            {
                return dispatcher.GetTypeInfo(0, 0);
            }
            catch (InvalidCastException exception)
            {
                throw new IDispatchNotImplementedException(exception);
            }
            catch
            {
                throw;
            }
        }
    }
}
