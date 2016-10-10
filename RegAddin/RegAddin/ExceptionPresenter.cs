using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class ExceptionPresenter
    {
        internal void ShowError(Exception exception)
        {
            Console.WriteLine("An unexpected erorr is occured.{2}{0} - {1}", exception.GetType().Name, exception.Message, Environment.NewLine);
        }
    }
}
