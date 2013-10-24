using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    public class AssemblyReference
    {
        public AssemblyReference()
        {
            Name = "";
            Path = "";
        }

        internal AssemblyReference(string name, string path)
        {
            Name = name;
            Path = path;
        }

        [DisplayName("Name"), Description("Name of the Assembly"), Category("Reference")]
        public string Name { get; set; }

        [DisplayName("Path"), Description("Full Path of the Assembly"), Category("Reference")]
        public string Path { get; set; }

        [DisplayName("IsExe"), Description("Gets info the reference is an executable"), Category("Reference")]
        public bool IsExe { get; set; }

        public override string ToString()
        {
            if (!String.IsNullOrWhiteSpace(Name))
                return Name;
            else
                return "Reference";
        }
    }
}
