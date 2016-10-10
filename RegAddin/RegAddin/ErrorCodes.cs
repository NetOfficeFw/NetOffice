using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class ErrorCodes : Dictionary<int, KeyValuePair<string, string>>
    {
        private static Exception _lastError;

        internal ErrorCodes()
        {
            Add(-1, new KeyValuePair<string, string>("NoArguments", "No Arguments Given."));
            Add(-2, new KeyValuePair<string, string>("MissingCommandOption", "Missing Root Command Option."));
            Add(-3, new KeyValuePair<string, string>("AmbiguousArguments", "More Than 1 Root Command Option Is Not Allowed."));
            Add(-4, new KeyValuePair<string, string>("AssemblyNotFound", "Invalid Assembly Path."));
            Add(-5, new KeyValuePair<string, string>("SyntaxError", "Syntax Error In 1 Or More Command Options."));
            Add(-6, new KeyValuePair<string, string>("InvalidCombination", "Illegal Command Options Combination."));
            Add(-7, new KeyValuePair<string, string>("UnkownArguments", "There Are 1 Or More Unkown Command Options Given."));
            Add(-8, new KeyValuePair<string, string>("InvalidArguments", "Invalid Arguments Has Been Given."));
            Add(-9, new KeyValuePair<string, string>("MissingArgument", "Missing Command Option Argument."));
            Add(-10, new KeyValuePair<string, string>("InvalidArgumentValue", "Unkown Command Option Argument."));
            Add(-11, new KeyValuePair<string, string>("UnauthorizedAccess", "Admin Permissions Required To Complete The Requested Transaction."));
            Add(-12, new KeyValuePair<string, string>("MissingPermissions", "Admin Permissions Required To Complete The Requested Transaction."));

            Add(-100, new KeyValuePair<string, string>("UnexpectedError", "Fatal/Unexpected Error."));
        }

        internal ErrorCodes SetLastError(Exception exception)
        {
            _lastError = exception;
            return this;
        }
    
        internal Exception GetLastError()
        {
            return _lastError;
        }

        internal string MessageFromCode(int code)
        {
            return this.First(e => e.Key == code).Value.Value;
        }

        internal string MessageDetailsFromCode(int code)
        {
            string result = this.First(e => e.Key == code).Value.Value;
            if (null != _lastError)
            { 
                result += Environment.NewLine + Environment.NewLine + "[" + _lastError.GetType().FullName +  "]";
                if(!String.IsNullOrWhiteSpace(_lastError.Message))
                {
                    result += Environment.NewLine;
                    result += "Error Code: " + _lastError.Message;
                }
            }
            return result;
        }

        internal int CodeFromName(string name)
        {
            return this.First(e => e.Value.Key.Equals(name, StringComparison.InvariantCultureIgnoreCase)).Key;
        }

        internal string MessageFromName(string name)
        {
           return this.First(e => e.Value.Key.Equals(name, StringComparison.InvariantCultureIgnoreCase)).Value.Value;
        }
    }
}
