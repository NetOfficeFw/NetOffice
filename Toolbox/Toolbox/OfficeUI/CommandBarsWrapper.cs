using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LateBindingApi.Core;

namespace NetOffice.DeveloperToolbox
{
    class CommandBarsWrapper : IEnumerable
    {
        COMObject _commandBars;

        public CommandBarsWrapper(COMObject innerObject)
        {
            _commandBars = innerObject;
        }

        private IEnumerator GetProxyEnumeratorAsProperty(COMObject comObject)
        {
            object enumProxy = Invoker.PropertyGet(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(comObject, enumProxy, true);
            Invoker.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                COMObject returnClass = new COMObject(enumerator, itemProxy);
                isMoveNextTrue = (bool)Invoker.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return returnClass;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return GetProxyEnumeratorAsProperty(_commandBars);
        }
    }
}
