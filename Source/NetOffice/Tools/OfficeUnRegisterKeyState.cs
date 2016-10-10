using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// RegAddin.exe want give the information that the necessary Office registry key(s) was already
    /// deleted from RegAddin or the addin need to do this by himself. (COMAddin base class want do this)
    /// </summary>
    [System.Runtime.InteropServices.Guid("DB8CD50A-2550-48A9-A124-7B65F94D7C36")]
    public enum OfficeUnRegisterKeyState
    {
        /// <summary>
        /// Registry key want be delete in the addin
        /// </summary>
        NeedToDelete = 0,

        /// <summary>
        /// Registry key is already deleted
        /// </summary>
        AlreadyDeleted = 1
    }
}
