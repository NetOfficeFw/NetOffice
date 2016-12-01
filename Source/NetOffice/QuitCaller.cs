using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    internal class QuitCaller
    {
        internal void TryCall(Settings settings, Invoker invoker, ICOMObject instance)
        {
            try
            {
                if (null == instance || null == instance.UnderlyingObject || null == settings || null == invoker)
                    return;
                if (settings.EnableAutomaticQuit)
                    invoker.Method(instance.UnderlyingObject, "Quit");
            }
            catch (Exception exception)
            {
                instance.Console.WriteException(exception);
            }
        }
    }
}
