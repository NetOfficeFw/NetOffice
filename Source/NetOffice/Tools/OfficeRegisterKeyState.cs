using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// RegAddin.exe want give the information that the necessary Office registry key(s) was already
    /// created from RegAddin or the addin need to do this by himself. (COMAddin base class want do this)
    /// </summary>
    [System.Runtime.InteropServices.Guid("D2AE173F-A763-4D70-86AA-5809DA5497AA")]
    public enum OfficeRegisterKeyState
    {
        /// <summary>
        /// Registry key want be create in the addin
        /// </summary>
        NeedToCreate = 0,

        /// <summary>
        /// Registry key is already created
        /// </summary>
        AlreadyCreated = 1
    }
}
