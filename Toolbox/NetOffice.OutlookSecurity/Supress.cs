using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.ComponentModel;
using System.Text;
using System.Timers;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace NetOffice.OutlookSecurity
{
    /// <summary>
    /// Outlook Security Suspressor
    /// </summary>
    public class Supress
    {
        #region Events

        public delegate void ErrorOccuredEventHandler(Exception exception);
        internal event ErrorOccuredEventHandler ErrorOccured;

        public delegate void SecurityDialogAction(SecurityDialog dialog, SecurityDialogCheckBox targetBox, SecurityDialogLeftButton targetButton);
        internal event SecurityDialogAction Action;

        public static event SecurityDialogAction OnAction
        {
            add 
            {
                if (null == _singleton)
                    _singleton = new Supress();
                _singleton.Action += value;
            }
            remove
            {
                if (null == _singleton)
                    _singleton = new Supress();
                _singleton.Action -= value;
            }
        }

        public static event ErrorOccuredEventHandler OnError
        {
            add
            {
                if (null == _singleton)
                    _singleton = new Supress();
                _singleton.ErrorOccured += value;
            }
            remove
            {
                if (null == _singleton)
                    _singleton = new Supress();
                _singleton.ErrorOccured -= value;
            }
        }

        #endregion

        #region Fields

        static Supress       _singleton;
        bool                 _enabled = false;
        System.Timers.Timer  _timer;
        StringBuilder        _strbClassName = new StringBuilder(255);
        StringBuilder        _strbCaption = new StringBuilder(255);
        List<SecurityDialog> _listDialogs = new List<SecurityDialog>();
        
        #endregion

        #region Construction

        private Supress()
        {
            _singleton = this;
            _timer = new System.Timers.Timer(500);
            _timer.Elapsed += new ElapsedEventHandler(_timer_Elapsed);

        }
         
        #endregion

        #region Properties

        /// <summary>
        /// Get or Set enabled state
        /// </summary>
        public static bool Enabled 
        {
            get 
            {
                if (null == _singleton)
                    _singleton = new Supress();
                return _singleton.EnabledState;
            }
            set
            {
                if (null == _singleton)
                    _singleton = new Supress();
                _singleton.EnabledState = value;
            }
        
        }
        
        internal bool EnabledState
        {
            get             
            {
                return _enabled;
            }
            set
            {
                if (_enabled == value)
                    return;

                _enabled = value;
                if (_enabled && !_timer.Enabled)
                    _timer.Enabled = true;
            }
        }

        #endregion

        #region Timer

        void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (!_enabled)
                return;

            try
            {
                TryGetOutlookSecurityDialogHandles();

                foreach (SecurityDialog item in _listDialogs)
                {
                    List<IntPtr> childWindows = User32.GetChildWindows(item.Handle);
                    if (!IncludeOneComboBox(childWindows))
                        continue;

                    IntPtr checkBox = GetCheckBox(childWindows, item.Handle);
                    string checkBoxText = User32.GetWindowText(checkBox);
                    if ((IntPtr.Zero != checkBox) && (!string.IsNullOrEmpty(checkBoxText)))
                        PostSendClick(checkBox);

                    IntPtr leftButton = GetLeftButton(childWindows, item.Handle);
                    string buttonText = User32.GetWindowText(leftButton);
                    if ((IntPtr.Zero != leftButton) && (!string.IsNullOrEmpty(buttonText)))
                        PostSendClick(leftButton);

                    if ((null != Action) && (null != checkBox) && (null != leftButton) && (checkBox != leftButton))
                    {
                        SecurityDialogCheckBox securityCheckbox = null;
                        SecurityDialogLeftButton securityButton = null;
                        if (null != checkBox)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(checkBox, out rect);
                            securityCheckbox = new SecurityDialogCheckBox(checkBox, User32.GetWindowText(checkBox), new Rect(rect));
                        }

                        if (null != leftButton)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(checkBox, out rect);
                            securityButton = new SecurityDialogLeftButton(leftButton, User32.GetWindowText(leftButton), new Rect(rect));
                        }

                        if ((!string.IsNullOrEmpty(checkBoxText)) && (!string.IsNullOrEmpty(buttonText)))
                            Action(item, securityCheckbox, securityButton);
                    }
                }
            }
            catch (Exception exception)
            {
                _enabled = false;
                if (null != ErrorOccured)
                    ErrorOccured(exception);
            }
        }

        #endregion

        #region Methods

        private void PostSendClick(IntPtr handle)
        {
            User32.RECT rect = new User32.RECT();
            User32.GetWindowRect(handle, out rect);
            User32.MouseClick(rect.Left + 15, rect.Top + 15);
            return;
        }

        private bool IncludeOneComboBox(List<IntPtr> childs)
        {
            int i = 0;
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if ("combobox" == className.ToLower())
                    i++;
            }

            return (1 == i);
        }

        private IntPtr GetCheckBox(List<IntPtr> childs, IntPtr dialogHandle)
        {
            IntPtr leftButton = IntPtr.Zero;
            User32.RECT rectFlag = new User32.RECT();
            rectFlag.Top = 100000;
            IntPtr leftChildButton = IntPtr.Zero;
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if ("Button" == className)
                { 
                    User32.RECT rect = new User32.RECT();
                    User32.GetWindowRect(item, out rect);
                    if (rect.Top < rectFlag.Top)
                    {
                        rectFlag.Top = rect.Top;
                        leftButton = item;
                    }
                }
            }
            return leftButton;
        }

        private IntPtr GetLeftButton(List<IntPtr> childs, IntPtr dialogHandle)
        {
            IntPtr leftButton = IntPtr.Zero;
            User32.RECT rectFlag = new User32.RECT();
            rectFlag.Left = 100000;
            IntPtr leftChildButton = IntPtr.Zero;
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if ("Button" == className)
                { 
                    User32.RECT rect = new User32.RECT();
                    User32.GetWindowRect(item, out rect);
                    if (rect.Left < rectFlag.Left)
                    {
                        rectFlag.Left = rect.Left;
                        leftButton = item;
                    }
                }
            }
            return leftButton;
        }

        private List<SecurityDialog> TryGetOutlookSecurityDialogHandles()
        {
            Process[] outlookProcess = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
            if (0 == outlookProcess.Length)
                return new List<SecurityDialog>();

            return GetSecurityDialogs(outlookProcess);
        }

        private List<SecurityDialog> GetSecurityDialogs(Process[] outlookProcess)
        {
            _listDialogs.Clear();

            User32.EnumDelegate filter = delegate(IntPtr hWnd, int lParam)
            {
                User32.GetClassName(hWnd, _strbClassName, _strbClassName.Capacity + 1);
                string className = _strbClassName.ToString();
               
                User32.GetWindowText(hWnd, _strbCaption, _strbCaption.Capacity + 1);
                string caption = _strbCaption.ToString();

                if(("#32770" == className.ToString()) && (User32.IsWindowVisible(hWnd)))
                {                   
                    uint processID = 0;
                    User32.GetWindowThreadProcessId(hWnd, out processID);
                    foreach (Process process in outlookProcess)
                    {
                        if (processID == process.Id)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(hWnd, out rect);
                           _listDialogs.Add(new SecurityDialog(hWnd, caption, className, new Rect(rect)));
                            break;
                        }
                    }
                }                
                return true;
            };

            User32.EnumDesktopWindows(IntPtr.Zero, filter, IntPtr.Zero);
            return _listDialogs;
        }

        #endregion
    }
}
