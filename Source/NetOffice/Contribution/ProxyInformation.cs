using NetOffice.ComTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Contribution
{
    /// <summary>
    /// Provides detailed information about a com proxy
    /// </summary>
    public class ProxyInformation
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name"></param>
        /// <param name="fullComponentName"></param>
        /// <param name="typeID"></param>
        public ProxyInformation(string name, string fullComponentName, Guid typeID)
        {
            Name = name;
            FullComponentName = fullComponentName;
            TypeID = typeID;
        }

        /// <summary>
        /// Proxy Interface/Class Name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Component Name
        /// </summary>
        public string FullComponentName { get; private set; }

        /// <summary>
        /// Interface/Class ID
        /// </summary>
        public Guid TypeID { get; private set; }

        /// <summary>
        /// Creates new instance of the class
        /// </summary>
        /// <param name="comProxy">target proxy</param>
        /// <returns>ProxyInformations instance</returns>
        /// <exception cref ="ArgumentNullException">argument is null</exception>
        public static ProxyInformation Create(object comProxy)
        {
            if (null == comProxy)
                throw new ArgumentNullException("comProxy");
            string className = TypeDescriptor.GetClassName(comProxy);
            string componentName = TypeDescriptor.GetComponentName(comProxy);
            Guid typeID = comProxy.TypeGuid();
            return new ProxyInformation(className, componentName, typeID);
        }
    }
}
