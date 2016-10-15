using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class ProgramHandler
    {
        internal void DoApplication(string[] args)
        {
            new VersionPresenter().ShowVersion();
            new CommandLineValidator().ValidateCommandLineArguments(ref args); 
            new CommandLineSettingsTransformer().ProceedCommandLineArguments(args);
            new OperationHandler().ProceedRequest(args);
        }
    }
}
