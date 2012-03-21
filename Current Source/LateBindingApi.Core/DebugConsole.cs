using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LateBindingApi.Core
{
    /// <summary>
    /// Operation mode for DebugConsole
    /// </summary>
    public enum ConsoleMode
    {
        /// <summary>
        /// debug log are not enabled
        /// </summary>
        None = 0,

        /// <summary>
        /// debug log was redirected to System.Console
        /// </summary>
        Console = 1,

        /// <summary>
        /// debug log append to a logfile
        /// </summary>
        LogFile = 2,

        /// <summary>
        /// hold all debug and exceptions logs in a internal string list
        /// </summary>
        MemoryList = 3
    }

    /// <summary>
    /// offers various debug, log and diagnostic functionality
    /// </summary>
    public static class DebugConsole
    {
        private static List<string> _messageList = new List<string>();

        /// <summary>
        /// append current time information in messages
        /// </summary>
        public static bool AppendTimeInfoEnabled { get; set; }

        /// <summary>
        /// operation mode
        /// </summary>
        public static ConsoleMode Mode { get; set; }

        /// <summary>
        /// name full file path and name of a logfile, must be set if Mode == LogFile
        /// </summary>
        public static string FileName { get; set; }

        /// <summary>
        /// returns all collected messages if Mode == MemoryList
        /// </summary>
        public static string[] Messages { get { return _messageList.ToArray(); } }

        /// <summary>
        /// clears message buffer
        /// </summary>
        public static void ClearMessagesList()
        {
            _messageList.Clear();
        }

        /// <summary>
        /// write log message
        /// </summary>
        /// <param name="message"></param>
        public static void WriteLine(string message)
        {
            string output = message;
            if (AppendTimeInfoEnabled)
                output = DateTime.Now.ToLongTimeString() + " - " + message;

            switch (Mode)
            {
                case ConsoleMode.Console:
                    Console.WriteLine(output);
                    break;
                case ConsoleMode.LogFile:
                    AppendToLogFile(output);
                    break;
                case ConsoleMode.MemoryList:
                    _messageList.Add(output);
                    break;
                case ConsoleMode.None:
                    // do nothing
                    break;
                default:
                    throw new ArgumentOutOfRangeException("Unkown Log Mode.");
            }
        }

        /// <summary>
        /// write exception log message
        /// </summary>
        /// <param name="exception"></param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static void WriteException(Exception exception)
        {
            string message = CreateExecptionLog(exception);
            WriteLine(message);
        }

        /// <summary>
        /// append message to logfile
        /// </summary>
        /// <param name="message"></param>
        private static void AppendToLogFile(string message)
        {
            if (null == FileName)
                throw new LateBindingApiException("FileName not set.");

            System.IO.File.AppendAllText(FileName, message + Environment.NewLine, Encoding.UTF8);
        }

        /// <summary>
        /// convert an exception to a string
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        private static string CreateExecptionLog(Exception exception)
        {
            string result = "";
            Exception ex = exception;
            while (ex != null)
            {
                string type = ex.GetType().Name;
                string message = ex.Message;
                string target = "<Empty>";
                if (null != ex.TargetSite)
                    target = ex.TargetSite.ToString();
                string trace = "<Empty>";
                if (null != ex.StackTrace)
                    trace = ex.StackTrace.ToString();

                result += "Type:" + type + Environment.NewLine;
                result += "Message:" + message + Environment.NewLine;
                result += "Target:" + target + Environment.NewLine;
                result += "Stack:" + trace + Environment.NewLine;

                result += Environment.NewLine;
                ex = ex.InnerException;
            }
            return result;
        }
    }
}
