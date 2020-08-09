﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.ComponentModel;
using System.Collections;

namespace NetOffice
{
    /// <summary>
    /// Offers various debug, log and diagnostic functionality
    /// </summary>
    public class DebugConsole : IEnumerable<DebugConsole.ConsoleMessage>
    {
        #region Nested

        /// <summary>
        /// Console Message Kind
        /// </summary>
        public enum MessageKind
        {
            /// <summary>
            /// Information Message
            /// </summary>
            Information = 0,

            /// <summary>
            /// Warning Message
            /// </summary>
            Warning = 1,

            /// <summary>
            /// Error Message
            /// </summary>
            Error = 2
        }

        /// <summary>
        /// Represents a debug console message
        /// </summary>
        public class ConsoleMessage
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="message">given message as any</param>
            public ConsoleMessage(string message)
            {
                Time = DateTime.Now;
                Message = message;
            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="kind">kind of given message</param>
            /// <param name="message">given message as any</param>
            public ConsoleMessage(MessageKind kind, string message)
            {
                Kind = kind;
                Time = DateTime.Now;
                Message = message;
            }

            /// <summary>
            /// Given Message
            /// </summary>
            public string Message { get; private set; }

            /// <summary>
            /// Message Time
            /// </summary>
            public DateTime Time { get; private set; }

            /// <summary>
            /// Kind of given message
            /// </summary>
            public MessageKind Kind { get; private set; }

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                return Message;
            }
        }

        /// <summary>
        /// Pipe Error Event Arguments
        /// </summary>
        public class PipeErrorEventArgs : EventArgs
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="pipeName">pipe name</param>
            /// <param name="text">message</param>
            /// <param name="error">error information</param>
            /// <param name="disableSharedOutput">disable shared output</param>
            public PipeErrorEventArgs(string pipeName, string text, Exception error, bool disableSharedOutput)
            {
                PipeName = pipeName;
                Text = text;
                Error = error;
                DisableSharedOutput = disableSharedOutput;
            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="pipeName">pipe name</param>
            /// <param name="text">message</param>
            /// <param name="error">error information</param>
            public PipeErrorEventArgs(string pipeName, string text, Exception error)
            {
                PipeName = pipeName;
                Text = text;
                Error = error;
                DisableSharedOutput = true;
            }

            /// <summary>
            /// Pipe Name
            /// </summary>
            public string PipeName { get; private set; }

            /// <summary>
            /// Message Text
            /// </summary>
            public string Text { get; private set; }

            /// <summary>
            /// Error Information
            /// </summary>
            public Exception Error { get; private set; }

            /// <summary>
            /// Disable the shared output
            /// </summary>
            public bool DisableSharedOutput { get; set; }
        }

        /// <summary>
        /// Message Added delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="message">new message</param>
        public delegate void MessageAddedHandler(DebugConsole sender, ConsoleMessage message);

        /// <summary>
        /// Message Removed delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="message">removed message</param>
        /// <param name="index">former message index</param>
        public delegate void MessageRemovedHandler(DebugConsole sender, ConsoleMessage message, int index);

        /// <summary>
        /// Message Clear delegate
        /// </summary>
        /// <param name="sender">sender instance</param>
        public delegate void MessageClearHandler(DebugConsole sender);

        #endregion

        #region Fields

        private static DebugConsole _default;

        private object _thisLock = new object();

        private static object _sharedLock = new object();
        
        private List<ConsoleMessage> _messageList = new List<ConsoleMessage>();

        private string _name = "";

        #endregion

        #region Events

        /// <summary>
        /// Occurs when a message has been added
        /// </summary>
        public event MessageAddedHandler MessageAdded;

        private void RaiseMessageAdded(ConsoleMessage message)
        {
            if (null != MessageAdded)
                MessageAdded(this, message);
        }

        /// <summary>
        ///  Occurs when a message has been removed
        /// </summary>
        public event MessageRemovedHandler MessageRemoved;

        private void RaiseMessageRemoved(ConsoleMessage message, int index)
        {
            if (null != MessageRemoved)
                MessageRemoved(this, message, index);
        }

        /// <summary>
        /// Occurs when the message list has been cleared
        /// </summary>
        public event MessageClearHandler MessageClear;

        private void RaiseMessageClear()
        {
            if (null != MessageClear)
                MessageClear(this);
        }

        /// <summary>
        /// Occurs when Console failed to send shared output
        /// </summary>
        public event EventHandler<PipeErrorEventArgs> PipeError;

        private bool RaisePipeError(string pipeName, string text, Exception error)
        {
            if (null != PipeError)
            {
                PipeErrorEventArgs args = new PipeErrorEventArgs(pipeName, text, error);
                PipeError(this, args);
                return args.DisableSharedOutput;                
            }
            else
                return true;
        }
        
        #endregion

        #region Properties

        /// <summary>
        /// Direct access to messages
        /// </summary>
        public IList<ConsoleMessage> MessagesInternal
        {
            get
            {
                return _messageList;
            }
        }

        /// <summary>
        /// Shared Default Instance
        /// </summary>
        public static DebugConsole Default
        {
            get
            {
                lock (_sharedLock)
                {
                    if (null == _default)
                        _default = new DebugConsole();
                    return _default;
                }

            }
        }
      
        /// <summary>
        /// Name of the Console instance
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                if (!String.IsNullOrEmpty(value))
                {
                    if (value.Length > 32)
                        throw new FormatException("Name length must be less than 32 characters");
                    char[] invalids = System.IO.Path.GetInvalidPathChars();
                    foreach (char item in invalids)
                    {
                        if (value.Contains(item.ToString()))
                            throw new FormatException("Name can't contain the '" + item.ToString() + "' character.");
                    }
                }
                _name = value;
            }
        }

        /// <summary>
        /// append current time information in messages
        /// </summary>
        public bool AppendTimeInfoEnabled { get; set; }

        /// <summary>
        /// operation mode
        /// </summary>
        public DebugConsoleMode Mode { get; set; }

        /// <summary>
        /// Send a all messages to a named pipe.
        /// </summary>
        public bool EnableSharedOutput { get; set; }

        /// <summary>
        /// Name full file path and name of a logfile, must be set if Mode == LogFile
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Returns all collected messages as a string enumerator
        /// </summary>
        public IEnumerable<string> Messages
        {
            get
            {
                List<string> list = new List<string>();
                foreach (ConsoleMessage item in _messageList)
                    list.Add(item.Message);
                return list;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Clears message buffer
        /// </summary>
        public void ClearMessagesList()
        {
            lock (_thisLock)
            {
                if (_messageList.Count > 0)
                {
                    _messageList.Clear();
                    RaiseMessageClear();
                }
            }
        }

        /// <summary>
        /// Write log message
        /// </summary>
        /// <param name="message">given message as any</param>
        /// <param name="args">message arguments</param>
        public void WriteLine(string message, params object[] args)
        {
            lock (_thisLock)
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

                if (AppendTimeInfoEnabled)
                    output = DateTime.Now.ToLongTimeString() + " - " + message;

                AddToMessageList(output, MessageKind.Information);

                switch (Mode)
                {
                    case DebugConsoleMode.Console:
                        Console.WriteLine(output);
                        break;
                    case DebugConsoleMode.Trace:
                        System.Diagnostics.Trace.WriteLine(output);
                        break;
                    case DebugConsoleMode.LogFile:
                        AppendToLogFile(output);
                        break;
                    case DebugConsoleMode.None:
                        // do nothing
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(Mode), @"Unknown value of DebugConsoleMode. Set the DebugConsole.Mode property to other value.");
                }

                TryWritePipe(output);
            }
        }

        /// <summary>
        /// Write log message
        /// </summary>
        /// <param name="any">given object as any</param>
        public void WriteLine(object any)
        {
            string message = null != any ? any.ToString() : "<null>";
            WriteLine(message);
        }

        /// <summary>
        /// Write log message
        /// </summary>
        /// <param name="message">given message as any</param>
        public void WriteLine(string message)
        {
            lock (_thisLock)
            {
                string output = message;
                
                if (AppendTimeInfoEnabled)
                    output = DateTime.Now.ToLongTimeString() + " - " + message;

                AddToMessageList(message, MessageKind.Information);

                switch (Mode)
                {
                    case DebugConsoleMode.Console:
                        Console.WriteLine(output);
                        break;
                    case DebugConsoleMode.Trace:
                        System.Diagnostics.Trace.WriteLine(output);
                        break;
                    case DebugConsoleMode.LogFile:
                        AppendToLogFile(output);
                        break;
                    case DebugConsoleMode.None:
                        // do nothing
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(Mode), @"Unknown value of DebugConsoleMode. Set the DebugConsole.Mode property to other value.");
                }

                TryWritePipe(output);
            }
        }

        /// <summary>
        /// Write exception log message
        /// </summary>
        /// <param name="exception"></param>
        public void WriteException(Exception exception)
        {
            lock (_thisLock)
            {
                string message = CreateExecptionLog(exception);
                AddToMessageList(message, MessageKind.Error);
                WriteLine(message);
            }          
        }
      
        /// <summary>
        /// Append message to logfile
        /// </summary>
        /// <param name="message"></param>
        private void AppendToLogFile(string message)
        {
            if (null == FileName)
                throw new NetOfficeException("Filename not set.");

            System.IO.File.AppendAllText(FileName, message + Environment.NewLine, Encoding.UTF8);
        }

        /// <summary>
        /// Convert an exception to a string
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        public static string CreateExecptionLog(Exception exception)
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

        /// <summary>
        /// Add a new item to the message list.
        /// </summary>
        /// <param name="text">text as any</param>
        private void AddToMessageList(string text)
        {
            AddToMessageList(text, MessageKind.Information);
        }

        /// <summary>
        /// Add a new item to the message list.
        /// </summary>
        /// <param name="text">text as any</param>
        /// <param name="kind">text kind</param>
        private void AddToMessageList(string text, MessageKind kind)
        {
            if (String.IsNullOrEmpty(text))
                return;

            ConsoleMessage message = new ConsoleMessage(kind, text);
            _messageList.Add(message);
            RaiseMessageAdded(message);

            if (_messageList.Count >= 110)
            {
                while (_messageList.Count >= 1000)
                {
                    ConsoleMessage deleteMessage = _messageList[0];
                    _messageList.RemoveAt(0);
                    RaiseMessageRemoved(deleteMessage, 0);
                }
            }
        }

        /// <summary>
        /// Try write message to named pipe
        /// </summary>
        /// <param name="text">text to send</param>
        public void TryWritePipe(string text)
        {
            if (!EnableSharedOutput)
                return;
            if (text == null || text == "")
                return;

            string name = "";
            if (Name == null || Name.Trim() == "")
                name = "NetOffice.DebugConsole";
            else
                name = "NetOffice.DebugConsole." + Name;

            try
            {
                using (System.IO.Pipes.NamedPipeClientStream pipe =
                    new System.IO.Pipes.NamedPipeClientStream(name))
                {
                    pipe.Connect(500);
                    using (StreamWriter writer = new StreamWriter(pipe))
                    {
                        writer.WriteLine(text);
                    }
                }
            }
            catch (TimeoutException exception)
            {
                if (RaisePipeError(name, text, exception))
                    EnableSharedOutput = false;
            }
            catch (Exception exception)
            {
                AddToMessageList("Failed to send shared message.", MessageKind.Warning);
                if (RaisePipeError(name, text, exception))
                    EnableSharedOutput = false;
            }         
        }

        #endregion

        #region IEnumerable<DebugConsole.ConsoleMessage>

        /// <summary>
        /// Returns an enumerable message sequence
        /// </summary>
        /// <returns>enumerator</returns>
        public IEnumerator<ConsoleMessage> GetEnumerator()
        {
            return _messageList.GetEnumerator();
        }

        /// <summary>
        ///  Returns an enumerable message sequence
        /// </summary>
        /// <returns>enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _messageList.GetEnumerator();
        }

        #endregion
    }
}
