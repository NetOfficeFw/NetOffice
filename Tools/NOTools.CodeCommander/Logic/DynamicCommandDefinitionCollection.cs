using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;
using NetOffice;
using NOTools.InMemoryCompiler;

namespace NOTools.CodeCommander.Logic
{
 
    class DynamicCommandDefinitionCollection : BindingList<DynamicCommandDefinition>
    {
        public void LoadFromFile(COMObject hostApplication)
        {
            return;
            DynamicCommandDefinition def1 = new DynamicCommandDefinition();
            def1.Definition.Name = "buhu1";
            def1.Definition.CustomClasses.AddNew(CommandBuilder.CreateCommandClass(hostApplication.ToString()));
            def1.Compile();
        }

        public void SaveToFile(COMObject hostApplication)
        { 
        }
    }
}
