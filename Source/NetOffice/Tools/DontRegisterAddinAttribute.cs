using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Don't register addin into the Microsoft Office application.
    /// The addin still will be registered as COM component from callers like RegAsm
    /// but it won't create or remove the Registry keys to bring the component into Office.
    /// (For troubleshooting purposes)
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class DontRegisterAddinAttribute : System.Attribute
    {

    }
}