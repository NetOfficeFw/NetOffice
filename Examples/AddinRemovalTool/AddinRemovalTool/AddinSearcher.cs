using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Linq;
using System.Text;

namespace AddinRemovalTool
{
    public enum ActionType
    {
        Start = 0,
        Error = 1
    }

    public delegate void ActionEventHandler(ActionType type, string action);
    
    class AddinSearcher : List<AddinEntry>
    {
        #region Properties

        string[] Products = new string[] { "Excel", "Word", "Outlook", "PowerPoint", "Access", "SuperAddin" };
        string[] Runtimes = new string[] { "2", "3", "35", "4", "45" };
        string[] Languages = new string[] { "CS", "VB" };
        string[] Names = new string[] { "SimpleAddin", "RibbonAddin", "TaskPaneAddin", "Addin" };

        string[] KnownAddins = new string[] { @"Software\Microsoft\Office\Excel\Addins\Sample.ExcelAddin",
                                              @"Software\Microsoft\Office\Word\Addins\Sample.WordAddin", 
                                              @"Software\Microsoft\Office\Outlook\Addins\Sample.TwitterAddin" };

        private string Actions { get; set; }

        private bool IsTransaction { get; set; }

        public event ActionEventHandler Action;

        internal void RaiseStartAction()
        {
            if (null != Action)
                Action(ActionType.Start, string.Empty);
        }

        internal void RaiseAction(ActionType type, string action)
        {
            if (IsTransaction)
            {
                Actions += action + Environment.NewLine;
            }
            else
            { 
                if(null != Action)
                    Action(type, action);
            }
        }

        #endregion

        public AddinSearcher()
        {
            Refresh();
        }

        internal void Refresh()
        {
            try
            {
                RaiseStartAction();
                this.Clear();
                FindStandardExampleAddins();
                FindWellKnownExampleAddins();
            }
            catch (Exception exception)
            {
                RaiseAction(ActionType.Error, exception.Message);
            }         
        }

        internal void BeginTransaction()
        {
            IsTransaction = true;
            Actions = string.Empty;
        }

        internal void EndTransaction()
        {
            IsTransaction = false;
            if (Actions != string.Empty)
            { 
                RaiseAction(ActionType.Error, Actions);
                Actions = string.Empty;
            }
        }

        private void FindStandardExampleAddins()
        {
            foreach (string product in Products)
            {
                foreach (string language in Languages)
                {
                    foreach (string runtime in Runtimes)
                    {
                        foreach (string name in Names)
                        {
                            string targetProgID = product + "Addin" + language + runtime + "." + name;
                            AddinEntry entry = KeyExists(product, targetProgID);
                            if (null != entry)
                                this.Add(entry);
                        }
                    }
                }
            }
        }

        private void FindWellKnownExampleAddins()
        {
            foreach (string item in KnownAddins)
            {
                string[] splitArray = item.Split(new string[] { @"\" }, StringSplitOptions.RemoveEmptyEntries);
                string product = splitArray[3];
                string targetProgID = splitArray[5];
                AddinEntry entry = KeyExists(product, targetProgID);
                if (null != entry)
                    this.Add(entry);
            }
        }
        
        private AddinEntry KeyExists(string product, string progID)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(String.Format(@"Software\Microsoft\Office\{0}\Addins\{1}", product, progID));
            if(null != key)
            {
                string friendlyName = key.GetValue("FriendlyName", string.Empty) as string;
                string description = key.GetValue("Description", string.Empty) as string;
                AddinEntry entry = new AddinEntry(this, product, progID, friendlyName, description);
                key.Close();
                return entry;
            }
            else
                return null;
        }
    }
}
