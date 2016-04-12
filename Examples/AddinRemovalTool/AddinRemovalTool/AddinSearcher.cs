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

        string[] Products = new string[] { "Excel", "Word", "Outlook", "PPoint", "Access" };
        string[] Runtimes = new string[] { "2", "3", "35", "4", "45" };
        string[] Languages = new string[] { "CS", "VB" };
        string[] Names = new string[] { "Simple", "Extended", "Tweak"};
        string[] KnownAddins = new string[] { @"Software\Microsoft\Office\Excel\Addins\NOSample.GoogleTranslation",
                                              @"Software\Microsoft\Office\Word\Addins\NOSample.Wikipedia", 
                                              @"Software\Microsoft\Office\Outlook\Addins\NOSample.Twitter" };
        string[] SuperAddins = new string[] { "NOToolsSuperAddin", "SuperAddin" };

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
                FindStandardExampleAddins();
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
                            string targetProgID = name + language + runtime + ".Addin";
                            AddinEntry entry = KeyExists(product, targetProgID);
                            if (null != entry)
                                this.Add(entry);
                        }                        
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
                            string targetProgID = name + product + language + runtime + ".Addin";
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

        #endregion
    }
}
