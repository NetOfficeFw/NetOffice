using System;
using System.Collections.Generic;
using NetOffice.Availity;

namespace NetOffice.Availity
{
    /// <summary>
    /// Performs a cache supported search to analyze at runtime a proxy supports a method or property
    /// </summary>
    internal class SupportedEntityFinder
    {
        /// <summary>
        /// Returns information a method or property is supported by a com proxy
        /// </summary>
        /// <param name="factory">core to perform searching</param>
        /// <param name="list">cache dictionary</param>
        /// <param name="searchType">entity type</param>
        /// <param name="proxy">target com proxy</param>
        /// <param name="name">name of the target entity</param>
        /// <returns>true if supported, otherwise false</returns>
        internal bool Find(Core factory, ref Dictionary<string,string> list, SupportedEntityType searchType, object proxy, string name)
        {
            switch (searchType)
            {
                case SupportedEntityType.Method:
                    {
                        if (null == list)
                        {
                            list = factory.GetSupportedEntities(proxy);
                            if (null == list)
                                return false;
                        }

                        string outValue = null;
                        return list.TryGetValue("Method-" + name, out outValue);
                    }
                case SupportedEntityType.Property:
                    {
                        if (null == list)
                        { 
                            list = factory.GetSupportedEntities(proxy);
                            if (null == list)
                                return false;
                        }

                        string outValue = null;
                        return list.TryGetValue("Property-" + name, out outValue);
                    }
                default:
                    {
                        if (null == list)
                        {
                            list = factory.GetSupportedEntities(proxy);
                            if (null == list)
                                return false;
                        }

                        string outValue = null;
                        bool result = list.TryGetValue("Property-" + name, out outValue);
                        if (result)
                            return true;

                        return list.TryGetValue("Method-" + name, out outValue);
                    }
            }

        }
    }
}
