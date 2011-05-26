using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace NetOffice.DeveloperUtils.ProcessKiller
{
    internal class OfficeProcessKiller : IDisposable 
    {
        #region Fields

        private NotifyIcon  _notify;
        private Icon        _runIcon;
        private Icon        _stopIcon;
        private Timer       _timer;
        private Keys        _key = Keys.A; 
        private Hotkey      _hotKey;
        private bool        _hotKeyEnabled;
        private Process[]   _excelProcs;
        private Process[]   _wordProcs;
        private Process[]   _outlookProcs;
        private Process[]   _powerProcs;
        private Process[]   _accessProcs;

        #endregion

        #region Construction

        internal OfficeProcessKiller(ListView listViewApps)
        {
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            _runIcon = new Icon(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".ProcessKiller.Running.ico"));
            _stopIcon = new Icon(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".ProcessKiller.NotRunning.ico"));
            _notify = new NotifyIcon();

            AttachedControl = listViewApps;          
            _timer = new Timer();
            _timer.Interval = 100;
            _timer.Tick += new EventHandler(_timer_Tick);
            _timer.Enabled = true;
        }

        #endregion

        #region Properties

        public bool TrayIcon
        {
            get
            {
                return _notify.Visible;
            }
            set
            {
                _notify.Visible = value;
            }
        }

        public ListView AttachedControl{get;set;}

        public Keys HotKey
        {
            get 
            {
                return _key; 
            }
            set
            {
                if (true == HotKeyEnabled)
                {
                    if (_hotKey != null)
                    { 
                        Hotkey.UnRegister(_hotKey);
                        _hotKey.Dispose();
                    }
                   _hotKey = Hotkey.Register(value);
                   _hotKey.HotkeyPressed += new EventHandler(_hotKey_HotkeyPressed);
                }
                _key = value;
            }
        }

        public bool HotKeyEnabled 
        {
            get
            {
                return _hotKeyEnabled;
            }
            set 
            {
                if (true == value)
                {
                    if (_hotKey != null)
                    {
                        Hotkey.UnRegister(_hotKey);
                        _hotKey.Dispose();
                    }
                    _hotKey = Hotkey.Register(_key);
                    _hotKey.HotkeyPressed += new EventHandler(_hotKey_HotkeyPressed);
                }
                else
                {
                    Hotkey.UnRegister(_hotKey);
                }
                _hotKeyEnabled = value;
            } 
        }

        public int WatchIntervallMs
        {
            get 
            {
                return _timer.Interval; 
            }
            set 
            {
                _timer.Interval = value; 
            }            
        }

        public bool WatchEnabled
        {
            get 
            {
                return _timer.Enabled;
            }
            set
            {
                _timer.Enabled = value;
            }
        }

        public bool Excel{get;set;}

        public bool Word { get; set; }

        public bool Outlook { get; set; }

        public bool PowerPoint { get; set; }

        public bool Access { get; set; }

        #endregion

        #region Methods

        public void KillProcesses()
        {
            if(Excel)
                KillProcesses(_excelProcs);
    
            if(Word)
                KillProcesses(_wordProcs);
            
            if(Outlook)
                KillProcesses(_outlookProcs);
            
            if(PowerPoint)
                KillProcesses(_powerProcs);

            if(Access)
                KillProcesses(_accessProcs);
        }

        private void KillProcesses(string name)
        {
            try
            {
                Process[] procs = Process.GetProcessesByName(name);

                foreach (Process p in procs)
                    p.Kill();
            }
            catch (System.ComponentModel.Win32Exception){;}
            catch (NotSupportedException){;}
            catch (InvalidOperationException){;}
        }

        private void ShowProcesses(string name, Process[] procs)
        {
            ListViewItem itemControl = null;
            foreach (ListViewItem item in AttachedControl.Items)
            {
                if (true == item.Text.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                {
                    itemControl = item;
                    break;
                }
            }

            if (null != itemControl)
            {
                string length = procs.Length.ToString();
                if(length != itemControl.SubItems[1].Text)
                    itemControl.SubItems[1].Text = length;
            }
        }

        private void KillProcesses(Process[] procs)
        {
            try
            {
                if (null == procs)
                    return;

                foreach (Process p in procs)
                    p.Kill();
            }
            catch
            {
                ; 
            }
        }

        private void ShowOfficeProcesses()
        {
            if (null != AttachedControl)
            { 
                Process[] procs = Process.GetProcessesByName("Excel");
                AttachedControl.Items[0].SubItems[1].Text = procs.Length.ToString();

                procs = Process.GetProcessesByName("WINWORD");
                AttachedControl.Items[1].SubItems[1].Text = procs.Length.ToString();

                procs = Process.GetProcessesByName("Outlook");
                AttachedControl.Items[2].SubItems[1].Text = procs.Length.ToString();
               
                procs = Process.GetProcessesByName("POWERPNT");
                AttachedControl.Items[3].SubItems[1].Text = procs.Length.ToString();

                procs = Process.GetProcessesByName("MSACCESS");
                AttachedControl.Items[4].SubItems[1].Text = procs.Length.ToString();
            }
        }
        
        private int ProcessCount()
        {
            int result = 0;

            if ((true == Excel) && (null != _excelProcs))
                result += _excelProcs.Length;

            if ((true == Word) && (null != _wordProcs))
                result += _wordProcs.Length;

            if ((true == Outlook) && (null != _outlookProcs))
                result += _outlookProcs.Length;

            if ((true == PowerPoint) && (null != _powerProcs))
                result += _powerProcs.Length;

            if ((true == Access) && (null != _accessProcs))
                result += _accessProcs.Length;

            return result;
        }

        #endregion

        #region Watch

        void _hotKey_HotkeyPressed(object sender, EventArgs e)
        {
            KillProcesses();
        }

        void _timer_Tick(object sender, EventArgs e)
        {
            _excelProcs = Process.GetProcessesByName("Excel");
            ShowProcesses("Excel", _excelProcs);

            _wordProcs = Process.GetProcessesByName("Winword");
            ShowProcesses("Winword", _wordProcs);

            _outlookProcs = Process.GetProcessesByName("Outlook");
            ShowProcesses("Outlook", _outlookProcs);

            _powerProcs = Process.GetProcessesByName("POWERPNT");
            ShowProcesses("POWERPNT", _powerProcs);

            _accessProcs = Process.GetProcessesByName("MSACCESS");
            ShowProcesses("MSACCESS", _accessProcs);

            int procCount = ProcessCount();
            if (procCount > 0)
            {
                _notify.Icon = _runIcon;
                _notify.Text = procCount.ToString() + " Office Instances";
            }
            else
            { 
                _notify.Icon = _stopIcon;
                _notify.Text = "";
            }
        }
        
        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (null != _hotKey)
            {
                _hotKey.Dispose();
                _hotKey = null;
            }

            if (null != _notify)
            {
                _notify.Visible = false;
                _notify.Dispose();
                _notify = null;
            }
        }

        #endregion
    }
}
