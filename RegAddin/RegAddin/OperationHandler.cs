using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class OperationHandler
    {
        internal void ProceedRequest(IEnumerable<string> args)
        {
            if (SingletonSettings.Diagnostics)
            {
                new Diag.Diagnostics().Show(args);
                return;
            }

            switch (SingletonSettings.Mode)
            {
                case SingletonSettings.ApplicationMode.Help:
                    new HelpPresenter().ShowHelp();
                    break;
                case SingletonSettings.ApplicationMode.Register:
                    new Register.RegisterOperationHandler().Proceed();    
                    break;
                case SingletonSettings.ApplicationMode.Unregister:
                    new Unregister.UnregisterOperationHandler().Proceed();
                    break;
                case SingletonSettings.ApplicationMode.RegFile:
                    new RegFile.RegFileOperationHandler().Proceed();
                    break;
                default:
                    throw new IndexOutOfRangeException();
            }
        }
    }
}
