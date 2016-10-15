using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    /// <summary>
    /// The well known entry point class
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Main entry point for the application
        /// </summary>
        /// <param name="args">Commandline arguments. see /help for documentation</param>
        /// <returns>0 if sucseed, otherwise a value below 0. see error codes in documentation</returns>
        public static int Main(string[] args)
        {            
            ConsoleAdapter.TryAttach();
            int returnCode = 0;

            try
            {
                new ProgramHandler().DoApplication(args);
            }
            catch (RegAddinException exception)
            {
                new ExceptionPresenter().ShowError(exception);
                returnCode = exception.ReturnCode;
            }
            catch (UnauthorizedAccessException exception)
            {
                new ExceptionPresenter().ShowError(exception);
                returnCode = new ErrorCodes().SetLastError(exception).CodeFromName("UnauthorizedAccess");
            }
            catch (System.Security.SecurityException exception)
            {
                new ExceptionPresenter().ShowError(exception);
                returnCode = new ErrorCodes().SetLastError(exception).CodeFromName("MissingPermissions");
            }
            catch (Exception exception)
            {
                new ExceptionPresenter().ShowError(exception);
                returnCode = new ErrorCodes().SetLastError(exception).CodeFromName("UnexpectedError");
            }
            
            if (SingletonSettings.Alert == SingletonSettings.AlertMode.On || 
                SingletonSettings.Alert == SingletonSettings.AlertMode.Error && returnCode != 0)
            {
                if (returnCode == 0)
                    Alert.Window.ShowSucceedMessage();
                else
                    Alert.Window.ShowError(new ErrorCodes().MessageDetailsFromCode(returnCode));
            }
            
            return returnCode;
        }
    }
}
