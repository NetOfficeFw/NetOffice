using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    public class SelectedAssembly
    {
        internal SelectedAssembly(string name, string path)
        {
            Name = name;
            Path = path;
        }

        public string Name { get; private set; }
        public string Path { get; private set; }
    }
}
