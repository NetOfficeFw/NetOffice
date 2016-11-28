using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    internal class EntityAvailableResolver
    {
        internal bool Resolve(Core factory, ref Dictionary<string,string> list, SupportEntityType searchType, object proxy, string name)
        {
            switch (searchType)
            {
                case SupportEntityType.Method:
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
                case SupportEntityType.Property:
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
