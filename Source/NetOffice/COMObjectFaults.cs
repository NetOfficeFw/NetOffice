using System;

namespace NetOffice
{
    /// <summary>
    /// Need to fix one design issue here in order to protect 
    /// the ICOMObject->ParentObject setter from public. (RemoveParent(ICOMObject comObject))
    /// COMDynamicObject prevent us to use a common internal base class because its alread inherites
    /// from System.Dynamic.DynamicObject.
    /// </summary>
    internal static class COMObjectFaults
    {
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

            // Todo: Check COMDuckObject here when its finished

            throw new ArgumentException("Unknown Instance Type " + comObject.GetType().FullName);
        }
    }
}
