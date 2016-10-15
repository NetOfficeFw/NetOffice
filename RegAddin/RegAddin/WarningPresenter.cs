using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class WarningPresenter
    {
        internal void ShowWarning(String message)
        {
            Console.WriteLine("Warning:{0}{1}", message, Environment.NewLine);
        }
    }
}
