using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;
using NetOffice;
using NOTools.InMemoryCompiler;

namespace NOTools.CodeCommander.Logic
{
    class DynamicCommandDefinition
    {
        public string Name
        {
            get
            {
                return "";
            }
            set { }
        }

        public bool IsReady
        {
            get
            {
                return true;
            }
        }

        public DynamicAssembly Definition { get; set; }
        public DynamicCommand Command { get; set; }

        public void Compile()
        {

        }

        public void ExecuteCommand()
        {
        }
    }

}
