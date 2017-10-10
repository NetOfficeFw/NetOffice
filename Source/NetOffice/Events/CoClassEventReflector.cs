using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace NetOffice.Events
{
    /// <summary>
    /// CoClass IEventBinding Services
    /// </summary>
    public static class CoClassEventReflector
    {
        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="type">target instance type</param> 
        /// <returns>true if one or more event is active, otherwise false</returns>
        public static bool HasEventRecipients(ICOMObject instance, Type type)
        {
            foreach (EventInfo item in type.GetEvents())
            {
                MulticastDelegate eventDelegate = (MulticastDelegate)type.GetType().GetField(item.Name,
                                                                            BindingFlags.NonPublic |
                                                                            BindingFlags.Instance).GetValue(instance);

                if ((null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0))
                    return false;
            }

            return false;
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="type">target instance type</param> 
        /// <param name="eventName">name of the event</param>
        /// <returns>true if event is active, otherwise false</returns>
        public static bool HasEventRecipients(ICOMObject instance, Type type, string eventName)
        {
            MulticastDelegate eventDelegate = (MulticastDelegate)type.GetField(
                                                "_" + eventName + "Event",
                                                BindingFlags.Instance |
                                                BindingFlags.NonPublic).GetValue(instance);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length > 0;
            }
            else
                return false;
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="type">target instance type</param>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <returns>actual event recipients</returns>
        public static Delegate[] GetEventRecipients(ICOMObject instance, Type type, string eventName)
        {
            MulticastDelegate eventDelegate = (MulticastDelegate)type.GetField(
                                                "_" + eventName + "Event",
                                                BindingFlags.Instance |
                                                BindingFlags.NonPublic).GetValue(instance);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates;
            }
            else
                return new Delegate[0];
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="type">target instance type</param>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <returns>count of event recipients</returns>
        public static int GetCountOfEventRecipients(ICOMObject instance, Type type, string eventName)
        {
            MulticastDelegate eventDelegate = (MulticastDelegate)type.GetField(
                                                "_" + eventName + "Event",
                                                BindingFlags.Instance |
                                                BindingFlags.NonPublic).GetValue(instance);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length;
            }
            else
                return 0;
        }

        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="instance">target instance</param>
        /// <param name="type">target instance type</param>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
        public static int RaiseCustomEvent(ICOMObject instance, Type type, string eventName, ref object[] paramsArray)
        {
            MulticastDelegate eventDelegate = (MulticastDelegate)type.GetField(
                                                "_" + eventName + "Event",
                                                BindingFlags.Instance |
                                                BindingFlags.NonPublic).GetValue(instance);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                foreach (var item in delegates)
                {
                    try
                    {
                        item.Method.Invoke(item.Target, paramsArray);
                    }
                    catch (Exception exception)
                    {
                        instance.Console.WriteException(exception);
                    }
                }

                if(instance.Settings.EnableAutoDisposeEventArguments)
                    Invoker.ReleaseParamsArray(paramsArray);
                return delegates.Length;
            }
            else
                return 0;
        }
    }
}