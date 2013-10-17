using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor.GACManagedAccess
{
    internal class GACAssembly
    {
        internal GACAssembly(string name, Version version, string path, string runtime, string pkToken)
        {
            Name = name;
            Version = version;
            Path = path;
            Runtime = runtime;
            PublicKeyToken = pkToken;
        }

        public string Name { get; private set; }

        public Version Version { get; private set; }

        public string Path { get; private set; }

        public string Runtime { get; private set; }

        public string PublicKeyToken { get; private set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
