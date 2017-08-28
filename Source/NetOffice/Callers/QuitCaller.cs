using System;

namespace NetOffice.Callers
{
    /// <summary>
    /// ICOMObject Quit method caller
    /// </summary>
    internal class QuitCaller
    {
        /// <summary>
        /// Try call "Quit" method without arguments when EnableAutomaticQuit is enabled
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>true if quit call proceed, otherwise false</returns>
        internal bool TryCall(ICOMObject instance)
        {
            if (null != instance)
                return TryCall(instance.Settings, instance.Invoker, instance);
            else
                return false;
        }

        /// <summary>
        /// Try call "Quit" method without arguments when EnableAutomaticQuit is enabled
        /// </summary>
        /// <param name="settings">settings to check EnableAutomaticQuit</param>
        /// <param name="invoker">invoker to call quit</param>
        /// <param name="instance">target instance</param>
        /// <returns>true if quit call proceed, otherwise false</returns>
        internal bool TryCall(Settings settings, Invoker invoker, ICOMObject instance)
        {
            try
            {
                if (null == instance || null == instance.UnderlyingObject || null == settings || null == invoker)
                    return false;
                if (settings.EnableAutomaticQuit)
                {
                    invoker.Method(instance.UnderlyingObject, "Quit");
                    return true;
                }
                else
                    return false;
            }
            catch (Exception exception)
            {
                instance.Console.WriteException(exception);
                return false;
            }
        }

        /// <summary>
        /// Try call "Quit" method without arguments even when EnableAutomaticQuit is disabled
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <returns>true if quit call proceed, otherwise false</returns>
        internal bool ForceTryCall(ICOMObject instance)
        {
            if (null != instance)
                return ForceTryCall(instance.Invoker, instance);
            else
                return false;
        }

        /// <summary>
        /// Try call "Quit" method without arguments even when EnableAutomaticQuit is disabled
        /// </summary>
        /// <param name="invoker">invoker to call quit</param>
        /// <param name="instance">target instance</param>
        /// <returns>true if quit call proceed, otherwise false</returns>
        internal bool ForceTryCall(Invoker invoker, ICOMObject instance)
        {
            try
            {
                if (null == instance || null == instance.UnderlyingObject || null == invoker)
                    return false;
                invoker.Method(instance.UnderlyingObject, "Quit");
                return true;
            }
            catch (Exception exception)
            {
                instance.Console.WriteException(exception);
                return false;
            }
        }
    }
}
