using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using NetOffice.NamedPipes;

namespace NetOffice
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
        MemoryList = 3,

        /// <summary>
        /// debug log was redirected to System.Diagnostics.Trace
        /// </summary>
        Trace = 4
    }

    /// <summary>
    /// Shared output connection technique
    /// </summary>
    public enum SharedOutputMode
    {
        /// <summary>
        /// IPC named pipes
        /// </summary>
        LocalNamedPipes = 0
    }

    /// <summary>
    /// offers various debug, log and diagnostic functionality
    /// </summary>
    public static class DebugConsole
    {
        private static object _sharedLock = new object();

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
        /// send a all messages to a named pipe. Use the NOTools.ConsoleMonitor to observe the console
        /// </summary>
        public static bool EnableSharedOutput { get; set; }

        /// <summary>
        /// Specify the shared output connection technique (currently ipc named pipes only. for future use to enable network and db logging)
        /// </summary>
        public static SharedOutputMode SharedOutputMode { get; set; }

        /// <summary>
        /// PipeConnection to NOTools.ConsoleMonitor
        /// </summary>
        private static PipeClient Pipe { get; set; }

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
        /// <param name="args"></param>
        public static void WriteLine(string message, params object[] args)
        {
            string output = message;

            int i = 0;
            foreach (object item in args)
            {
                string replaceValue = "";
                if (null != item)
                    replaceValue = item.ToString();
                output = output.Replace("{" + i.ToString() + "}", replaceValue);
                i++;
            }

            if (ConsoleMode.Console == Mode || ConsoleMode.Trace == Mode)
                output = "NetOffice: " + output;

            if (AppendTimeInfoEnabled)
                output = DateTime.Now.ToLongTimeString() + " - " + message;

            switch (Mode)
            {
                case ConsoleMode.Console:
                    Console.WriteLine(output);
                    break;
                case ConsoleMode.Trace:
                    System.Diagnostics.Trace.WriteLine(output);
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

            InternalSendNamedPipeMessage(output, null);
        }

        /// <summary>
        /// write log message
        /// </summary>
        /// <param name="message"></param>
        public static void WriteLine(string message)
        {
            string output = message;

            if (ConsoleMode.Console == Mode || ConsoleMode.Trace == Mode)
                output = "NetOffice: " + output;

            if (AppendTimeInfoEnabled)
                output = DateTime.Now.ToLongTimeString() + " - " + message;

            switch (Mode)
            {
                case ConsoleMode.Console:
                    Console.WriteLine(output);
                    break;
                case ConsoleMode.Trace:
                    System.Diagnostics.Trace.WriteLine(output);
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

            InternalSendNamedPipeMessage(output, null);
        }

        /// <summary>
        /// write exception log message
        /// </summary>
        /// <param name="exception"></param>
        public static void WriteException(Exception exception)
        {
            string message = CreateExecptionLog(exception);
            WriteLine(message);
        }

        /// <summary>
        /// Send a message to the NOTools.Console monitor pipe
        /// </summary>
        /// <param name="console">name for the console(must exclude the '?' char) or null for default console</param>
        /// <param name="message">the given message as any</param>
        /// <returns>entry id for the log message if arrived, otherwise null</returns>
        public static string SendPipeConsoleMessage(string console, string message)
        {
            try
            {
                lock (_sharedLock)
                {
                    if (null == Pipe)
                        Pipe = new PipeClient();
                    return Pipe.SendConsoleMessage(console, message, null);
                }
            }
            catch (Exception exception)
            {
                EnableSharedOutput = false;
                WriteException(exception);
                return null;
            }
        }
    
        /// <summary>
        /// Send a message to the NOTools.Console monitor pipe
        /// </summary>
        /// <param name="console">name for the console(must exclude the '?' char) or null for default console</param>
        /// <param name="message">the given message as any</param>
        /// <param name="parentEntryID">parent message id. the console monitor can show a hierarchy with these info</param>
        /// <returns>entry id for the log message if arrived, otherwise null</returns>
        public static string SendPipeConsoleMessage(string console, string message, string parentEntryID)
        {
            try
            {
                lock (_sharedLock)
                {
                    if (null == Pipe)
                        Pipe = new PipeClient();
                    return Pipe.SendConsoleMessage(console, message, parentEntryID);
                }
            }
            catch (Exception exception)
            {
                EnableSharedOutput = false;
                WriteException(exception);
                return null;
            }
        }

        /// <summary>
        /// Send a channel message to the NOTools.Console monitor pipe
        /// </summary>
        /// <param name="channel">channel id string. the argument must exclude the '?' character</param>
        /// <param name="message">the given message as any</param>
        /// <returns>entry id for the log message if arrived, otherwise null</returns>
        public static string SendPipeChannelMessage(string channel, string message)
        {
            try
            {
                lock (_sharedLock)
                {
                    if (null == Pipe)
                        Pipe = new PipeClient();
                    return Pipe.SendChannelMessage(channel, message);
                }
            }
            catch (Exception exception)
            {
                EnableSharedOutput = false;
                WriteException(exception);
                return null;
            }
        }


        /// Send a message to the NOTools.Console monitor pipe
        /// </summary>
        /// <param name="message">the given message as any</param>
        /// <returns>true if send</returns>
        internal static string InternalSendNamedPipeMessage(string message, string parentEntryID)
        {
            try
            {
                if (!EnableSharedOutput)
                    return null;
                lock (_sharedLock)
                {
                    if (null == Pipe)
                        Pipe = new PipeClient();
                    return Pipe.SendConsoleMessage(null, message, parentEntryID);
                }
            }
            catch (Exception exception)
            {
                EnableSharedOutput = false;
                WriteException(exception);
                return null;
            }
        }

        /// <summary>
        /// Send a channel message to the NOTools.Console monitor pipe
        /// </summary>
        /// <param name="channel">channel id string. the argument must exclude the '?' character</param>
        /// <param name="message">the given message as any</param>
        /// <returns>true if send</returns>
        internal static string InternalSendNamedPipeChannelMessage(string channel, string message)
        {
            try
            {
                if (!EnableSharedOutput)
                    return null;
                lock (_sharedLock)
                {
                    if (null == Pipe)
                        Pipe = new PipeClient();
                    return Pipe.SendChannelMessage(channel, message);
                }
            }
            catch (Exception exception)
            {
                EnableSharedOutput = false;
                WriteException(exception);
                return null;
            }
        }

        /// <summary>
        /// append message to logfile
        /// </summary>
        /// <param name="message"></param>
        private static void AppendToLogFile(string message)
        {
            if (null == FileName)
                throw new NetOfficeException("FileName not set.");

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
