using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using NetOffice.Exceptions;

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
        /// <exception cref="AvailityException">An unexpected error occurs. See inner exception(s) for details.</exception>
        internal bool Find(Core factory, ref Dictionary<string,string> list, SupportedEntityType searchType, object proxy, string name)
        {
            try
            {
                switch (searchType)
                {
                    case SupportedEntityType.Method:
                        {
                            if (null == list)
                            {
                                list = GetSupportedEntities(factory, proxy);
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
                                list = GetSupportedEntities(factory, proxy);
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
                                list = GetSupportedEntities(factory, proxy);
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
            catch (Exception exception)
            {
                throw new AvailityException(exception);
            }            
        }

        /// <summary>
        /// Creates an entity support list for a proxy
        /// </summary>
        /// <param name="factory">core to perform searching</param>
        /// <param name="comProxy">proxy to analyze</param>
        /// <returns>supported methods and properties as name/kind dictionary</returns>
        /// <exception cref="COMException">Throws generaly if any exception occurs. See inner exception(s) for details</exception>
        internal Dictionary<string, string> GetSupportedEntities(Core factory, object comProxy)
        {
            try
            {                
                Guid parentLibraryGuid = CoreFactoryExtensions.GetParentLibraryGuid(factory, comProxy);
                if (Guid.Empty == parentLibraryGuid)
                    return null;

                string className = TypeDescriptor.GetClassName(comProxy);
                string key = (parentLibraryGuid.ToString() + className).ToLower();

                Dictionary<string, string> supportList = null;
                if (factory.EntitiesListCache.TryGetValue(key, out supportList))
                    return supportList;

                supportList = new Dictionary<string, string>();
                IDispatch dispatch = comProxy as IDispatch;
                if (null == dispatch)
                    throw new COMException("Unable to cast underlying proxy to IDispatch.");

                COMTypes.ITypeInfo typeInfo = dispatch.GetTypeInfo(0, 0);
                if (null == typeInfo)
                    throw new COMException("GetTypeInfo returns null.");

                IntPtr typeAttrPointer = IntPtr.Zero;
                typeInfo.GetTypeAttr(out typeAttrPointer);

                COMTypes.TYPEATTR typeAttr = (COMTypes.TYPEATTR)Marshal.PtrToStructure(typeAttrPointer, typeof(COMTypes.TYPEATTR));
                for (int i = 0; i < typeAttr.cFuncs; i++)
                {
                    string strName, strDocString, strHelpFile;
                    int dwHelpContext;
                    IntPtr funcDescPointer = IntPtr.Zero;
                    COMTypes.FUNCDESC funcDesc;
                    typeInfo.GetFuncDesc(i, out funcDescPointer);
                    funcDesc = (COMTypes.FUNCDESC)Marshal.PtrToStructure(funcDescPointer, typeof(COMTypes.FUNCDESC));

                    switch (funcDesc.invkind)
                    {
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYGET:
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUT:
                        case COMTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF:
                            {
                                typeInfo.GetDocumentation(funcDesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                                string outValue = "";
                                bool exists = supportList.TryGetValue("Property-" + strName, out outValue);
                                if (!exists)
                                    supportList.Add("Property-" + strName, strDocString);
                                break;
                            }
                        case COMTypes.INVOKEKIND.INVOKE_FUNC:
                            {
                                typeInfo.GetDocumentation(funcDesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                                string outValue = "";
                                bool exists = supportList.TryGetValue("Method-" + strName, out outValue);
                                if (!exists)
                                    supportList.Add("Method-" + strName, strDocString);
                                break;
                            }
                    }

                    typeInfo.ReleaseFuncDesc(funcDescPointer);
                }

                typeInfo.ReleaseTypeAttr(typeAttrPointer);
                Marshal.ReleaseComObject(typeInfo);

                factory.EntitiesListCache.Add(key, supportList);

                return supportList;
            }
            catch (Exception exception)
            {
                throw new COMException("An unexpected error occurs.", exception);
            }
        }
    }
}
