using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Duck
{
    internal class DuckTypeIssueClassGenerator : IDisposable
    {     
        internal DuckTypeIssueClassGenerator(StringBuilder classBuilder, DuckInterface proxyInterface, string implementationName)
        {
            Builder = classBuilder;
            ImplementationName = implementationName + "_";
            ProxyInterfaceFullName = proxyInterface.FullName;
            SyntaxClassRequired = proxyInterface.SyntaxClassRequired;

            if (!SyntaxClassRequired)
                return;

            Builder.AppendLine(@"using NetRuntimeSystem = System;");
            Builder.AppendLine(@"using System;");
            Builder.AppendLine(@"using System.Collections.Generic;");
            Builder.AppendLine(@"using NetOffice;");
            Builder.AppendLine(@"namespace NetOffice.DuckTyping");
            Builder.AppendLine("{");

            Builder.AppendLine("\tpublic class " + ImplementationName + " : COMObject");
            Builder.AppendLine("\t{");
            

            Builder.AppendLine("\t\t#region Ctor");
            Builder.AppendLine("");
            Builder.AppendLine("\t\tpublic " + ImplementationName + "(Core factory, ICOMObject parentObject, object proxy) : base(factory, parentObject, proxy)\r\n\t\t{");

            Builder.AppendLine("\t\t}");
            Builder.AppendLine("");
            Builder.AppendLine("\t\t#endregion");
        }

        public StringBuilder Builder { get; private set; }

        public string ImplementationName { get; private set; }

        public string ProxyInterfaceFullName { get; private set; }

        private bool SyntaxClassRequired { get; set; }

        public void Dispose()
        {
            if (!SyntaxClassRequired)
                return;

            Builder.AppendLine("\t}");
            Builder.AppendLine("}");
        }
    }
}
