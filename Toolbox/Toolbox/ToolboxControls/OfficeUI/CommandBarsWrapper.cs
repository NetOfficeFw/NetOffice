using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeUI
{
    /// <summary>
    /// Wrapper for the office application commandbars
    /// </summary>
    internal class CommandBarsWrapper : IEnumerable
    {
        #region Fields

        private COMObject _commandBars;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerObject">commandbars proxy from the application</param>
        public CommandBarsWrapper(COMObject innerObject)
        {
            _commandBars = innerObject;
        }

        #endregion

        #region Methods

        private IEnumerator GetProxyEnumeratorAsProperty(COMObject comObject)
        {
            object enumProxy = Invoker.Default.PropertyGet(comObject, "_NewEnum");
            COMObject enumerator = new COMObject(Core.Default, comObject, enumProxy, true);
            Invoker.Default.MethodWithoutSafeMode(enumerator, "Reset", null);
            bool isMoveNextTrue = (bool)Invoker.Default.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.Default.PropertyGetWithoutSafeMode(enumerator, "Current", null);
                COMObject returnClass = new COMObject(enumerator, itemProxy);
                isMoveNextTrue = (bool)Invoker.Default.MethodReturnWithoutSafeMode(enumerator, "MoveNext", null);
                yield return returnClass;
            }
        }

        #endregion

        #region IEnumerable

        public IEnumerator GetEnumerator()
        {
            return GetProxyEnumeratorAsProperty(_commandBars);
        }

        #endregion
    }
}
