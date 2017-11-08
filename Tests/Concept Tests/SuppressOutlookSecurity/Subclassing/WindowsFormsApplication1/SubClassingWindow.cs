using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsFormsApplication1
{
    /// <summary>
    /// Event handler for SubClassingWindow MessageRecieved event
    /// </summary>
    /// <param name="handle">target window handle</param>
    /// <param name="m">recived message</param>
    public delegate void MessageRecievedEventHandler(IntPtr handle, WndMessage m);

    /// <summary>
    /// Filter mode for SubClassingWindow Filter
    /// </summary>
    public enum MessageFilterMode
    { 
        /// <summary>
        /// recieved message type must be specified in the filter
        /// </summary>
        WhiteList = 0,

        /// <summary>
        /// recieved message type must be excluded in the filter
        /// </summary>
        BlackList = 1
    }

    /// <summary>
    /// 
    /// </summary>
    internal class SubClassingWindow : System.Windows.Forms.NativeWindow
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        public SubClassingWindow(IntPtr handle)
        {
            Filter = new WndMessage[0];
            AssignHandle(handle);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        /// <param name="messageRecieved">function pointer for the message recieved event</param>
        public SubClassingWindow(IntPtr handle, MessageRecievedEventHandler messageRecieved)
        {
            Filter = new WndMessage[0];
            AssignHandle(handle);
            MessageRecieved = messageRecieved;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        /// <param name="callBack">function pointer for the message recieved event</param>
        /// <param name="messageRecieved">filter to customize(filter) the message recieve event</param>
        public SubClassingWindow(IntPtr handle, MessageRecievedEventHandler messageRecieved, WndMessage[] filter)
        { Filter = new WndMessage[0];
            AssignHandle(handle);
            MessageRecieved = messageRecieved;
            Filter = filter;
        }

        #endregion

        #region Events

        /// <summary>
        /// Occures when the target window has recieved a message
        /// </summary>
        public event MessageRecievedEventHandler MessageRecieved;

        /// <summary>
        /// Raise the MessageRecieved event
        /// </summary>
        /// <param name="handle">target window handle</param>
        /// <param name="m">recivied message</param>
        private void RaiseMessageRecieved(IntPtr handle, WndMessage m)
        {
            if (null != MessageRecieved)
                MessageRecieved(handle, m);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Additional message filter for MessageRecieved event 
        /// </summary>
        public WndMessage[] Filter { get; set; }

        /// <summary>
        /// Message filter mode
        /// </summary>
        public MessageFilterMode FilterMode { get; set; }

        #endregion

        #region Overrides
        
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            base.WndProc(ref m);
            WndMessage wndMessage = WndMessage.WM_NULL;

            if (null != Filter && Filter.Length != 0)
            {
                System.Enum.TryParse<WndMessage>(m.Msg.ToString(), out wndMessage);
                if (WndMessage.WM_NULL == wndMessage)
                    return;

                if (FilterMode == MessageFilterMode.WhiteList)
                {
                    foreach (var item in Filter)
                    {
                        if(wndMessage == item)
                        {
                            RaiseMessageRecieved(Handle, wndMessage);
                            return;
                        }
                    }                  
                }
                else
                {
                    bool found = false;
                    foreach (var item in Filter)
                    {
                        if (wndMessage == item)
                        {
                            found = true;
                            break;
                        }
                    }

                    if(true == found)
                        RaiseMessageRecieved(Handle, wndMessage);
                }
            }

           
        }
        
        #endregion
    }
}
