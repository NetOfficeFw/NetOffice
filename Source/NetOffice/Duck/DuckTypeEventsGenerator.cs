using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NetOffice.Duck
{
    internal class DuckTypeEventsGenerator : IDisposable
    {
        internal DuckTypeEventsGenerator(StringBuilder builder, EventInfo[] events)
        {
            HasEvents = null != events && events.Length > 0;
            if (!HasEvents)
                return;

            Builder = builder;

            Builder.AppendLine(Environment.NewLine + "\t\t#region IEventBinding" + Environment.NewLine + Environment.NewLine);
            Builder.AppendLine(Resources.EventBinding);
            Builder.AppendLine(Environment.NewLine + "\t\t#endregion");
            Builder.AppendLine(Environment.NewLine + "\t\t#region Events" + Environment.NewLine + Environment.NewLine);

            foreach (EventInfo item in events)
            {
                if (item.EventHandlerType.GetGenericArguments().Length > 0)
                    continue;
                    
                ParameterInfo[] eventArgs = item.EventHandlerType.GetMethod("Invoke").GetParameters();
                foreach (var arg in eventArgs)
                {
                    string fieldName = "_" + item.Name.Substring(0,1).ToLower() + item.Name.Substring(1);
                    string eventProperty = "\t\tpublic event " + item.EventHandlerType.FullName + " " + item.Name + Environment.NewLine + "\t\t{" + Environment.NewLine;
                    eventProperty += "\t\t\tadd" + Environment.NewLine;
                    eventProperty += "\t\t\t{" + Environment.NewLine;
                    eventProperty += "\t\t\t\tCreateEventBridge();" + Environment.NewLine;
                    eventProperty += "\t\t\t\t" + fieldName + " += value;" + Environment.NewLine;
                    eventProperty += "\t\t\t}" + Environment.NewLine;

                    eventProperty += "\t\t\tremove" + Environment.NewLine;
                    eventProperty += "\t\t\t{" + Environment.NewLine;
                    eventProperty += "\t\t\t\t" + fieldName + " -= value;" + Environment.NewLine;
                    eventProperty += "\t\t\t}" + Environment.NewLine;
                    eventProperty += "\t\t}" + Environment.NewLine;
                    eventProperty += "\t\t" + "private " + item.EventHandlerType.FullName + " " + fieldName + ";" + Environment.NewLine + Environment.NewLine;
                    
                    Builder.Append(eventProperty);
                }
            }
        }
        
        private StringBuilder Builder { get; set; }

        private bool HasEvents { get; set; }

        public void Dispose()
        {
            if (HasEvents)
                Builder.AppendLine(Environment.NewLine + "\t\t#endregion");
        }
    }
}
