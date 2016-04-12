using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Encapsulate generic addin services 
    /// </summary>
    public abstract class COMAddinBase
    {
        /// <summary>
        /// Host Application Instance
        /// </summary>
        public abstract COMObject AppInstance { get; }
    }
}
