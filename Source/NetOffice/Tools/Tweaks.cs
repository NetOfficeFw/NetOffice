#define RegKeyDisposeAvailable

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;
using System.Text;
using NetOffice.Tools;
using NetOffice.Filtering;

namespace NetOffice.Tools
{
    /// <summary>
    /// Points to an addin method that try to detect the addin is loaded from HKEY_LOCAL_MACHINE\Software\Office or HKEY_CURRENT_USER\Software\Office
    /// Each COMAddin base class has a coresponding -cache supported- method for this delegate.
    /// COMAddin base class want give this method as delegate during the loading process to service methods.
    /// The service methods want call the delegate only if need because it is potentialy expensive in performance
    /// which is a problem in a loading process.
    /// </summary>
    /// <returns>null if unknown or true/false</returns>
    public delegate bool? IsLoadedFromSystemKeyDelegate();
}
