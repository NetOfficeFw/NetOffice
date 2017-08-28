using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Duck
{
    internal class DuckTypeClassGenerator : IDisposable
    {
        internal DuckTypeClassGenerator(StringBuilder classBuilder, DuckInterface proxyInterface, string implementationName, string issueImplementationName)
        {
            Builder = classBuilder;
            ImplementationName = implementationName;
            ProxyInterfaceFullName = proxyInterface.FullName;

            bool eventsSupported = proxyInterface.IsValidEventClass;

            Builder.AppendLine(@"using NetRuntimeSystem = System;");
            Builder.AppendLine(@"using System;");
            Builder.AppendLine(@"using System.Collections;");
            Builder.AppendLine(@"using System.Collections.Generic;");
            Builder.AppendLine(@"using NetOffice;");
            Builder.AppendLine(@"namespace NetOffice.DuckTyping");
            Builder.AppendLine("{");

            Builder.AppendLine("\tpublic class " + ImplementationName + " : " + (proxyInterface.SyntaxClassRequired ? issueImplementationName : "COMObject") + ", " + ProxyInterfaceFullName);
            Builder.AppendLine("\t{");

            if (eventsSupported)
            {
                Builder.AppendLine("\t\t#region Fields");
                Builder.AppendLine("");
                Builder.AppendLine("\t\tprivate static Type _thisType;");
                Builder.AppendLine("\t\tprivate KeyValuePair<string, Type>[] _sinks;");
                Builder.AppendLine("\t\tprivate NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;");
                Builder.AppendLine("\t\tprivate SinkHelper _activeSink;");
                Builder.AppendLine("\t\tprivate string _activeSinkId;" + Environment.NewLine);
                Builder.AppendLine("\t\t#endregion" + Environment.NewLine);
            }
            else
            {
                Builder.AppendLine("\t\t#region Fields");
                Builder.AppendLine("");
                Builder.AppendLine("\t\tprivate KeyValuePair<string, Type>[] _sinks;");
                Builder.AppendLine("\t\tprivate static Type _thisType;" + Environment.NewLine);
                Builder.AppendLine("\t\t#endregion" + Environment.NewLine);
            }

            Builder.AppendLine("\t\t#region Ctor");
            Builder.AppendLine("");
            Builder.AppendLine("\t\tpublic " + ImplementationName + "(Core factory, ICOMObject parentObject, object proxy) : base(factory, parentObject, proxy)\r\n\t\t{");

            if (eventsSupported)
            {
                Builder.Append("\t\t\t_sinks = new KeyValuePair<string, Type>[]{");
                KeyValuePair<string, Type>[] sinks = proxyInterface.EventSinks;
                for (int i = 0; i < sinks.Length; i++)
                {
                    Builder.Append("new KeyValuePair<string, Type>(" + "\"" + sinks[i].Key + "\"" + ", typeof(" + sinks[i].Value.FullName + "))");
                    if (i < sinks.Length - 1)
                        Builder.Append(",");
                }
                Builder.AppendLine("};");
            }
            else
            {
                Builder.AppendLine("\t\t\t_sinks = new KeyValuePair<string, Type>[0];");
            }

            Builder.AppendLine("\t\t}");
            Builder.AppendLine("");
            Builder.AppendLine("\t\t#endregion");

            Builder.AppendLine(Environment.NewLine + "\t\t#region Type Information" + Environment.NewLine);

            Builder.AppendLine("\t\tpublic static Type LateBindingApiWrapperType");
            Builder.AppendLine("\t\t{");
            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");
            Builder.AppendLine("\t\t\t\tif (null == _thisType)");
            Builder.AppendLine("\t\t\t\t\t_thisType = typeof("+ implementationName + ");");
            Builder.AppendLine("\t\t\t\t return _thisType;");
            Builder.AppendLine("\t\t\t}");
            Builder.AppendLine("\t\t}" + Environment.NewLine);

            Builder.AppendLine("\t\tpublic override Type InstanceType");
            Builder.AppendLine("\t\t{");
            Builder.AppendLine("\t\t\tget");
            Builder.AppendLine("\t\t\t{");
            Builder.AppendLine("\t\t\t\treturn LateBindingApiWrapperType;");
            Builder.AppendLine("\t\t\t}");
            Builder.AppendLine("\t\t}" + Environment.NewLine);

            //public static Type LateBindingApiWrapperType

            Builder.AppendLine("\t\t#endregion");
        }

        public StringBuilder Builder { get; private set; }

        public string ImplementationName { get; private set; }

        public string ProxyInterfaceFullName { get; private set; }
        
        public void Dispose()
        {
            Builder.AppendLine("\t}");
            Builder.AppendLine("}");
        }
    }
}
