using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    internal static class COMObjectFaults
    {
        /*
            RemoveParent(ICOMObject comObject):
            Need to make one design fail here in order to protect 
            the ICOMObject->ParentObject setter from public.
            (COMDynamicObject prevent us to use an internal base class because its alread inherites
            from System.Dynamic.DynamicObject)
        */
        /// <summary>
        /// Remove parent object from ICOMObject instance
        /// </summary>
        /// <param name="comObject">target instance</param>
        internal static void RemoveParent(ICOMObject comObject)
        {
            COMObject instance1 = comObject as COMObject;
            if (null != instance1)
            {
                instance1.ParentObject = null;
                return;
            }

            COMDynamicObject instance2 = comObject as COMDynamicObject;
            if (null != instance2)
            {
                instance1.ParentObject = null;
                return;
            }
            
            throw new ArgumentException("Unknown Instance Type " + comObject.GetType().FullName);
        }
    }
}
