using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AddinRemovalTool
{
    class AddinEntry
    {
        public AddinEntry(AddinSearcher parent, string product, string progID, string name, string description)
        {
            Parent = parent;
            Product = product;
            ProgID = progID;
            Name = name;
            Description = description;
        }

        private AddinSearcher Parent { get; set; }
        public string Product { get; private set; }
        public string ProgID { get; private set; }
        public string Name { get; private set; }
        public string Description { get; private set; }

        public bool Delete()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKey(String.Format(@"Software\Microsoft\Office\{0}\Addins\{1}", Product, ProgID), true);
                return true;
            }
            catch (Exception exception)
            {
                Parent.RaiseAction(ActionType.Error, exception.Message);
                return false;
            }
        }
    }
}
