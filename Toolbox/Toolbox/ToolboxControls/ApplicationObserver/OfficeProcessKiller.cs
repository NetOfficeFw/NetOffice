using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ApplicationObserver
{
    /// <summary>
    /// process watcher for ms-office products
    /// </summary>
    internal class OfficeApplicationObserver : IDisposable
    {
        #region Fields

        private string      _killQuestion = "Ausgewählte Instanzen löschen?";
        private bool        _showQuesionBeforeKill;
        private NotifyIcon  _notify;
        private Icon        _runIcon;
        private Icon        _stopIcon;
        private Timer       _timer;
        private Keys        _key = Keys.A;
        private Hotkey      _hotKey;
        private bool        _hotKeyEnabled;
        private Process[]   _allProcs = new Process[0];
        private Process[]   _excelProcs;
        private Process[]   _wordProcs;
        private Process[]   _outlookProcs;
        private Process[]   _powerProcs;
        private Process[]   _accessProcs;
        private Process[]   _projectProcs;
        private Process[]   _visioProcs;
        private int         _currentLanguageID = 1031;

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="listViewApps">view control</param>
        internal OfficeApplicationObserver(ListView listViewApps)
        {
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            _runIcon = new Icon(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".ToolboxControls.ApplicationObserver.IconsAndConfig.Running.ico"));
            _stopIcon = new Icon(this.GetType().Assembly.GetManifestResourceStream(assemblyName + ".ToolboxControls.ApplicationObserver.IconsAndConfig.NotRunning.ico"));
            _notify = new NotifyIcon();

            AttachedControl = listViewApps;
            _timer = new Timer();
            _timer.Interval = 100;
            _timer.Tick += new EventHandler(Timer_Tick);
            _timer.Enabled = true;
        }

        #endregion

        #region Events

        /// <summary>
        /// The whole process list has been changed
        /// </summary>
        public event EventHandler AllProcessesChanged;

        /// <summary>
        /// an office process has been started or terminated
        /// </summary>
        public event EventHandler InstanceRunningCountChanged;

        #endregion

        #region Properties

        /// <summary>
        /// Current user language id
        /// </summary>
        public int CurrentLanguageID
        {
            get
            {
                return _currentLanguageID;
            }
            set
            {
                _currentLanguageID = value;
            }
        }

        /// <summary>
        /// Question for the user before kill someone
        /// </summary>
        public string KillQuestion
        {
            get
            {
                return _killQuestion;
            }
            set
            {
                _killQuestion = value;
            }
        }

        /// <summary>
        /// Show KillQuestion as message box before kill
        /// </summary>
        public bool ShowQuestionBeforeKill
        {
            get
            {
                return _showQuesionBeforeKill;
            }
            set
            {
                _showQuesionBeforeKill = value;
            }
        }

        /// <summary>
        /// Show tray icon to indicates one ore more (selected) office application is running
        /// </summary>
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

        /// <summary>
        /// View control to show current processes in real time
        /// </summary>
        public ListView AttachedControl{get;set;}

        /// <summary>
        /// Hotkey definition to kill instances with one hit
        /// </summary>
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
                   _hotKey.HotkeyPressed += new EventHandler(HotKey_HotkeyPressed);
                }
                _key = value;
            }
        }

        /// <summary>
        /// Hotkey definition is enabled
        /// </summary>
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
                    _hotKey.HotkeyPressed += new EventHandler(HotKey_HotkeyPressed);
                }
                else
                {
                    Hotkey.UnRegister(_hotKey);
                }
                _hotKeyEnabled = value;
            }
        }

        /// <summary>
        /// Refresh intervall for the process list
        /// </summary>
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

        /// <summary>
        /// Enable to observe the process list
        /// </summary>
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

        /// <summary>
        /// Enable to kill excel
        /// </summary>
        public bool Excel{get;set;}

        /// <summary>
        /// Enable to kill word
        /// </summary>
        public bool Word { get; set; }

        /// <summary>
        /// Enable to kill outlook
        /// </summary>
        public bool Outlook { get; set; }

        /// <summary>
        /// Enable to kill power point
        /// </summary>
        public bool PowerPoint { get; set; }

        /// <summary>
        /// Enable to kill access
        /// </summary>
        public bool Access { get; set; }

        /// <summary>
        /// Enable to kill project
        /// </summary>
        public bool Project { get; set; }

        public bool Visio { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Kill selected processes if running
        /// </summary>
        public void KillProcesses()
        {
            if(DialogResult.Yes != MessageBox.Show(_killQuestion, "NetOffice Developer Toolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                return ;

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

            if (Project)
                KillProcesses(_projectProcs);

            if (Visio)
                KillProcesses(_visioProcs);

        }

        private void KillProcesses(string name)
        {
            try
            {
                Process[] procs = Process.GetProcessesByName(name);

                foreach (Process p in procs)
                    p.Kill();
            }
            catch (System.ComponentModel.Win32Exception) { ;}
            catch (NotSupportedException) { ;}
            catch (InvalidOperationException) { ;}
            catch(Exception)
            {
                throw;
            }
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
                if (length != itemControl.SubItems[1].Text)
                {
                    itemControl.SubItems[1].Text = length;
                    if (null != InstanceRunningCountChanged)
                        InstanceRunningCountChanged(this, new EventArgs());
                }
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

                procs = Process.GetProcessesByName("WINPROJ");
                AttachedControl.Items[5].SubItems[1].Text = procs.Length.ToString();

                procs = Process.GetProcessesByName("VISIO");
                AttachedControl.Items[6].SubItems[1].Text = procs.Length.ToString();
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

            if ((true == Project) && (null != _projectProcs))
                result += _projectProcs.Length;

            if ((true == Visio) && (null != _visioProcs))
                result += _visioProcs.Length;

            return result;
        }

        private static bool IsOfficeProcess(Process process)
        {
            string name = process.ProcessName.ToUpper();
            switch (name)
            {
                case "EXCEL":
                case "WINWORD":
                case "OUTLOOK":
                case "POWERPNT":
                case "MSACCESS":
                case "WINPROJ":
                case "VISIO":
                    return true;
                default:
                    return false;
            }
        }

        private static Process[] SortProcesses(Process[] allNewProcs)
        {
            List<Process> resultList = new List<Process>();
            foreach (Process item in allNewProcs)
            {
                if (IsOfficeProcess(item))
                    resultList.Insert(0, item);
                else
                    resultList.Add(item);
            }

            return resultList.ToArray();
        }

        #endregion

        #region Watch

        private void HotKey_HotkeyPressed(object sender, EventArgs e)
        {
            KillProcesses();
        }

        private void CheckChangedProcs(Process[] allNewProcs)
        {
            if (allNewProcs.Length != _allProcs.Length)
            {
                if (null != AllProcessesChanged)
                {
                    allNewProcs = SortProcesses(allNewProcs);
                    AllProcessesChanged(allNewProcs, new EventArgs());
                }
            }
            else
            {
                // check some new
                foreach (Process newProcess in allNewProcs)
                {
                    bool found = false;
                    foreach (Process item in _allProcs)
                    {
                        if (item.Id == newProcess.Id)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        if (null != AllProcessesChanged)
                        {
                            allNewProcs = SortProcesses(allNewProcs);
                            AllProcessesChanged(allNewProcs, new EventArgs());
                        }
                        return;
                    }
                }

                // check deleted process
                foreach (Process oldProcess in _allProcs)
                {
                    bool found = false;
                    foreach (Process newProcess in allNewProcs)
                    {
                        if (newProcess.Id == oldProcess.Id)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        if (null != AllProcessesChanged)
                            AllProcessesChanged(allNewProcs, new EventArgs());
                        return;
                    }
                }
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                Process[] allProcs = Process.GetProcesses();
                CheckChangedProcs(allProcs);
                _allProcs = allProcs;

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

                _projectProcs = Process.GetProcessesByName("WINPROJ");
                ShowProcesses("WINPROJ", _projectProcs);

                _visioProcs = Process.GetProcessesByName("VISIO");
                ShowProcesses("VISIO", _visioProcs);

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
            catch (Exception exception)
            {
                _timer.Enabled = false;
                Forms.ErrorForm.ShowError(null, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion

        #region IDisposable

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
