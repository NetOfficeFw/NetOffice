using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// Mark a static method as registry export handler.
    /// RegAddin.exe want call this method when an .reg file export is requested.
    /// The method can give additional registry informations to the export.
    /// The method need the following signature: static RegExport CreateRegistryExport(InstallScope scope);
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Method, AllowMultiple = false)]
    public class RegExportFunctionAttribute : System.Attribute
    {
        
    }
}
