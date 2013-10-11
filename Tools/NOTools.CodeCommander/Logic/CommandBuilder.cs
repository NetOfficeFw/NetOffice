using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.DeveloperAddin.Logic
{
    internal class CommandBuilder
    {
        internal static string CreateCommandClass(string officeApplication)
        {
            switch (officeApplication)
            {
                default:
                    break;
            }
            return "";
        }

        private static string CommandTemplate
        {
            get 
            {
                if (null == _commandTemplate)
                { 
                }
                return _commandTemplate;
            }
        }
        private static string _commandTemplate;
    }
}
