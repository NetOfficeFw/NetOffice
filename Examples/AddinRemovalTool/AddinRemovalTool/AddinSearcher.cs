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
    
    internal class AddinSearcher : List<AddinEntry>
    {
        #region Ctor

        public AddinSearcher()
        {
            Refresh();
        }

        #endregion

        #region Properties / Events

        string[] plains = new string[] { "SimpleAddin", "RibbonAddin", "TaskPaneAddin" };
        string[] Products = new string[] { "Excel", "Word", "Outlook", "PowerPoint", "Access" };
        string[] Runtimes = new string[] { "2", "3", "35", "4", "45" };
        string[] Languages = new string[] { "CS", "VB" };
        string[] Names = new string[] { "01Addin", "02Addin", "03Addin", "04Addin", "05Addin", "06Addin", "07Addin", "08Addin" };
        string[] KnownAddins = new string[] { @"Software\Microsoft\Office\Excel\Addins\NOSample.GoogleTranslation",
                                              @"Software\Microsoft\Office\Word\Addins\NOSample.Wikipedia", 
                                              @"Software\Microsoft\Office\Outlook\Addins\NOSample.Twitter" };
        string[] SuperAddins = new string[] { "SuperAddin", "Super2Addin" };

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
                if (null != Action)
                    Action(type, action);
            }
        }

        #endregion

        #region Methods

        internal void Refresh()
        {
            try
            {
                RaiseStartAction();
                this.Clear();
                FindPlainAddins();
                FindStandardExampleAddins();
                FindLoaderShimExampleAddins();
                FindWellKnownExampleAddins();
                FindSuperAddins();
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

        private void FindPlainAddins()
        {
            foreach (string product in Products)
            {
                foreach (string language in Languages)
                {
                    foreach (var runtime in Runtimes)
                    {
                        foreach (string item in plains)
                        {
                            string targetProgId = product + "Addin" + language + runtime + "." + item;
                            AddinEntry entry = KeyExists(product, targetProgId);
                            if (null != entry)
                                this.Add(entry);
                        }
                    }                 
                }
            }
        }

        private void FindSuperAddins()
        {
            foreach (string name in SuperAddins)
            {
                foreach (string language in Languages)
                {
                    foreach (string runtime in Runtimes)
                    {
                        foreach (string product in Products)
                        {
                            string targetProgID = name + language + runtime + ".Connect";
                            AddinEntry entry = KeyExists(product, targetProgID);
                            if (null != entry)
                                this.Add(entry);
                        }                        
                    }
                }            
            }
        }

        private void FindLoaderShimExampleAddins()
        {
            foreach (string product in Products)
            {
                foreach (string language in Languages)
                {
                    foreach (string runtime in Runtimes)
                    {
                        string targetProgID = "LoaderShim" + language + runtime + ".Connect";
                        AddinEntry entry = KeyExists(product, targetProgID);
                        if (null != entry)
                            this.Add(entry);
                    }
                }
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
                            string targetProgID = product + name + language + runtime + ".Connect";
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
            if (product == "PPoint") // bad hot fix - please dont look!
                product = "PowerPoint";

            RegistryKey key = Registry.CurrentUser.OpenSubKey(String.Format(@"Software\Microsoft\Office\{0}\Addins\{1}", product, progID));
            if (null != key)
            {
                string friendlyName = key.GetValue("FriendlyName", string.Empty) as string;
                string description = key.GetValue("Description", string.Empty) as string;
                AddinEntry entry = new AddinEntry(this, product, progID, friendlyName, description, false);
                key.Close();
                return entry;
            }
            else
            {
                key = Registry.LocalMachine.OpenSubKey(String.Format(@"Software\Microsoft\Office\{0}\Addins\{1}", product, progID));
                if (null != key)
                {
                    string friendlyName = key.GetValue("FriendlyName", string.Empty) as string;
                    string description = key.GetValue("Description", string.Empty) as string;
                    AddinEntry entry = new AddinEntry(this, product, progID, friendlyName, description, true);
                    key.Close();
                    return entry;
                }
            }

            return null;
        }

        #endregion
    }
}
