using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;
using System.Timers;
using System.Diagnostics;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
{
    /// <summary>
    /// Outlook Security Automation
    /// </summary>
    public class Automation : IDisposable 
    {
        #region Fields

        private static string _timeOutMessage = "Unable to suppress security dialog {0}.";

        private static string _outlookProcessName = "OUTLOOK";
        private static string _dialogClassName = "#32770";
        private static string _buttonClassName = "button";
        private static string _checkBoxClassName = "button";
        private static string _comboClassName = "combobox";
        private static string _progressClassName = "msctls_progress32";
        private static string _richClassName = "richedit20wpt";
        private static int _timerIntervalMS = 500;
        private static int _delaySeconds = 5;

        private event SecurityDialogAction _onAction;
        private ErrorOccuredEventHandler _onError;

        private object _lock = new object();
        private StringBuilder _strbClassName = new StringBuilder(255);
        private StringBuilder _strbCaption = new StringBuilder(255);
        private Dictionary<SecurityDialog, DateTime> _listDialogs = new Dictionary<SecurityDialog, DateTime>();
        private System.Timers.Timer _timer;
        private int _timeOutSeconds = 10;
        private ClickStrategy _strategy;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Automation()
        {
            _timer = new System.Timers.Timer(_timerIntervalMS);
            _timer.Elapsed += new ElapsedEventHandler(Timer_Elapsed);
        }

        #endregion

        #region Events

        /// <summary>
        /// OnError event handler
        /// </summary>
        /// <param name="exception">exception as any</param>
        public delegate void ErrorOccuredEventHandler(System.Exception exception);

        /// <summary>
        /// Dialog Popup event handler
        /// </summary>
        /// <param name="dialog">shown dialog</param>
        /// <param name="targetBox">checbox in dialog</param>
        /// <param name="targetButton">Allow button in dialog</param>
        public delegate void SecurityDialogAction(SecurityDialog dialog, SecurityDialogCheckBox targetBox, SecurityDialogLeftButton targetButton);

        /// <summary>
        /// Occurs when the security dialog is shown
        /// </summary>
        public event SecurityDialogAction OnAction
        {
            add
            {
                _onAction += value;
            }
            remove
            {
                _onAction -= value;
            }
        }

        /// <summary>
        /// Occurs when an error occured in automation
        /// </summary>
        public event ErrorOccuredEventHandler OnError
        {
            add
            {
                _onError += value;
            }
            remove
            {
                _onError -= value;
            }
        }

        #endregion
        
        #region Properties

        /// <summary>
        /// Get or set enabled state. Default:false
        /// </summary>
        public bool Enabled
        {
            get
            {
                return _timer.Enabled;
            }
            set
            {
                lock (_lock)
                {
                    if (_timer.Enabled != value)
                        _timer.Enabled = value;
                }
            }
        }

        /// <summary>
        /// If a security dialog is still visible after TimeOutSeconds 
        /// the OnError event occurs and Suppress stop try close the dialog. Default:10
        /// </summary>
        public int TimeOutSeconds
        {
            get
            {
                return _timeOutSeconds;
            }
            set
            {
                lock (_lock)
                {
                    if (_timeOutSeconds != value)
                        _timeOutSeconds = value;
                }
            }
        }

        /// <summary>
        /// Current used strategy to simulate click for security dialogs 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public ClickStrategy Strategy
        {
            get
            {
                return _strategy;
            }
            set
            {
                lock (_lock)
                {
                    if (_strategy != value)
                        _strategy = value;
                }
            }
        }

        #endregion

        #region Methods

        private void ProceedSuppress()
        {
            lock (_lock)
            {
                TryGetOutlookSecurityDialogHandles();

                foreach (KeyValuePair<SecurityDialog, DateTime> item in _listDialogs)
                {
                    List<IntPtr> childWindows = User32.GetChildWindows(item.Key.Handle);
                    if (!ContainsOneComboBoxOrRichEdit(childWindows))
                        continue;

                    if ((DateTime.Now - item.Value).TotalSeconds >= _timeOutSeconds)
                    {
                        if (!item.Key.ExceptionThrown)
                        {
                            if (null != _onError)
                                _onError(new TimeoutException(String.Format(_timeOutMessage, item.Key.Handle)));
                            item.Key.ExceptionThrown = true;
                        }
                        continue;
                    }

                    IntPtr checkBox = GetCheckBox(childWindows, item.Key.Handle);
                    string checkBoxText = User32.GetWindowText(checkBox);
                    if ((IntPtr.Zero != checkBox) && (!string.IsNullOrEmpty(checkBoxText)))
                    {
                        if (!item.Key.CheckBoxPassed)
                            DoControlClick(checkBox);
                        item.Key.CheckBoxPassed = true;
                    }

                    IntPtr leftButton = GetLeftButton(childWindows, item.Key.Handle);
                    string buttonText = User32.GetWindowText(leftButton);
                    if ((IntPtr.Zero != leftButton) && (!string.IsNullOrEmpty(buttonText)))
                    {
                        if (ContainsOneVisibleProgress(childWindows, item.Key.Handle) & ((DateTime.Now - item.Value).TotalSeconds <= _delaySeconds))
                            continue;
                        EnableControl(leftButton);
                        DoControlClick(leftButton);
                    }

                    if ((null != _onAction) && (null != checkBox) && (IntPtr.Zero != leftButton) && (checkBox != leftButton))
                    {
                        SecurityDialogCheckBox securityCheckbox = null;
                        SecurityDialogLeftButton securityButton = null;
                        if (null != checkBox)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(checkBox, out rect);
                            securityCheckbox = new SecurityDialogCheckBox(checkBox, User32.GetWindowText(checkBox), new Rect(rect));
                        }

                        if (IntPtr.Zero != leftButton)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(checkBox, out rect);
                            securityButton = new SecurityDialogLeftButton(leftButton, User32.GetWindowText(leftButton), new Rect(rect));
                        }

                        if ((!string.IsNullOrEmpty(checkBoxText)) && (!string.IsNullOrEmpty(buttonText)))
                            _onAction(item.Key, securityCheckbox, securityButton);
                    }
                }
            }
        }

        private static void EnableControl(IntPtr handle)
        {
            User32.EnableWindow(handle, true);
        }

        private void DoControlClick(IntPtr handle)
        {
            switch (_strategy)
            {
                case ClickStrategy.MoveTo:
                    DoMouseMoveClick(handle);
                    break;
                case ClickStrategy.SendMessage:
                    DoSendClick(handle);
                    break;
                case ClickStrategy.PostMessage:
                    DoPostClick(handle);
                    break;
                default:
                    break;
            }
        }

        private static void DoSendClick(IntPtr handle)
        {
            User32.DoSendMouseClick(handle);
        }

        private static void DoPostClick(IntPtr handle)
        {
            User32.DoPostMouseClick(handle);
        }

        private static void DoMouseMoveClick(IntPtr handle)
        {
            User32.RECT rect = new User32.RECT();
            User32.GetWindowRect(handle, out rect);
            User32.DoMouseMoveClick(rect.Left + 15, rect.Top + 15);
        }

        private static bool ContainsOneComboBoxOrRichEdit(List<IntPtr> childs)
        {
            int comboCount = 0;
            int richCount = 0;

            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item).ToLower();
                if (_comboClassName == className)
                    comboCount++;
                else if (_richClassName == className)
                    richCount++;
            }

            return (1 == comboCount || 1 == richCount);
        }

        private static bool ContainsOneVisibleProgress(List<IntPtr> childs, IntPtr dialogHandle)
        {
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if (_progressClassName == className.ToLower() && User32.IsWindowVisible(item))
                    return true;
            }
            return false;
        }

        private static IntPtr GetCheckBox(List<IntPtr> childs, IntPtr dialogHandle)
        {
            IntPtr leftButton = IntPtr.Zero;
            User32.RECT rectFlag = new User32.RECT();
            rectFlag.Top = 100000;
            IntPtr leftChildButton = IntPtr.Zero;
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if (_checkBoxClassName == className.ToLower())
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

        private static IntPtr GetLeftButton(List<IntPtr> childs, IntPtr dialogHandle)
        {
            IntPtr leftButton = IntPtr.Zero;
            User32.RECT rectFlag = new User32.RECT();
            rectFlag.Left = 100000;
            IntPtr leftChildButton = IntPtr.Zero;
            foreach (IntPtr item in childs)
            {
                string className = User32.GetClassName(item);
                if (_buttonClassName == className.ToLower())
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

        private Dictionary<SecurityDialog, DateTime> TryGetOutlookSecurityDialogHandles()
        {
            Process[] outlookProcess = System.Diagnostics.Process.GetProcessesByName(_outlookProcessName);
            if (0 == outlookProcess.Length)
                return new Dictionary<SecurityDialog, DateTime>();

            return GetSecurityDialogs(outlookProcess);
        }

        private Dictionary<SecurityDialog, DateTime> GetSecurityDialogs(Process[] outlookProcess)
        {
            Dictionary<SecurityDialog, DateTime> local = new Dictionary<SecurityDialog, DateTime>();

            User32.EnumDelegate filter = delegate (IntPtr hWnd, int lParam)
            {
                User32.GetClassName(hWnd, _strbClassName, _strbClassName.Capacity + 1);
                string className = _strbClassName.ToString();

                User32.GetWindowText(hWnd, _strbCaption, _strbCaption.Capacity + 1);
                string caption = _strbCaption.ToString();

                if ((_dialogClassName == className.ToString()) && (User32.IsWindowVisible(hWnd)))
                {
                    uint processID = 0;
                    User32.GetWindowThreadProcessId(hWnd, out processID);
                    foreach (Process process in outlookProcess)
                    {
                        if (processID == process.Id)
                        {
                            User32.RECT rect = new User32.RECT();
                            User32.GetWindowRect(hWnd, out rect);
                            local.Add(new SecurityDialog(hWnd, caption, className, new Rect(rect)), DateTime.Now);
                            break;
                        }
                    }
                }
                return true;
            };
            User32.EnumDesktopWindows(IntPtr.Zero, filter, IntPtr.Zero);

            foreach (KeyValuePair<SecurityDialog, DateTime> item in local)
            {
                SecurityDialog existing = GetDialogFromHandle(_listDialogs, item.Key.Handle);
                if (null == existing)
                    _listDialogs.Add(item.Key, item.Value);
            }

            List<SecurityDialog> toDelete = new List<SecurityDialog>();
            foreach (KeyValuePair<SecurityDialog, DateTime> item in _listDialogs)
            {
                SecurityDialog existing = GetDialogFromHandle(_listDialogs, item.Key.Handle);
                if (null == existing)
                    toDelete.Add(item.Key);
            }
            foreach (SecurityDialog item in toDelete)
                _listDialogs.Remove(item);

            return _listDialogs;
        }

        private static SecurityDialog GetDialogFromHandle(Dictionary<SecurityDialog, DateTime> dictionary, IntPtr handle)
        {
            foreach (KeyValuePair<SecurityDialog, DateTime> item in dictionary)
            {
                if (item.Key.Handle == handle)
                    return item.Key;
            }
            return null;
        }

        #endregion

        #region IDisposable

        /// <summary>
        /// Cleanup the instance
        /// </summary>
        public void Dispose()
        {
            lock (_lock)
            {
                if (null != _timer)
                {
                    _timer.Dispose();
                    _timer = null;
                }
            }
        }

        #endregion

        #region Timer

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (!Enabled)
                return;

            try
            {
                ProceedSuppress();
            }
            catch (System.Exception exception)
            {
                Enabled = false;
                if (null != _onError)
                    _onError(exception);
            }
        }

        #endregion
    }
}
