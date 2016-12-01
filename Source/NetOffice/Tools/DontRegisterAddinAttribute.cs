using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Dont register addin into the office application.
    /// The addin still want be registered as COM component from callers like RegAsm
    /// but dont create/remove the Registry keys to bring the component into Office.
    /// (For troubleshooting purpose) 
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class DontRegisterAddinAttribute : System.Attribute
    {

    }
}