using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Handle all display Item messages und organize them in plain and hierarchy
    /// </summary>
    internal class ConsoleViewItemCollection
    {
        #region Fields

        private static string _spaceString = "   ";
        private static string _emptyMachineDefault = "<Unkown>";
        private static string _emptyAppDomainDefault = "<Undefined>";
        private static int _stringBuilderDefaultSize = 2048;

        private int _idCounter = 0;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ConsoleViewItemCollection() 
        {
            Plain = new List<ConsoleViewItem>();
            Hierarchy = new List<ConsoleViewItem>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Plain item collection
        /// </summary>
        public List<ConsoleViewItem> Plain { get; private set; }

        /// <summary>
        /// Hierarchy item collection
        /// </summary>
        public List<ConsoleViewItem> Hierarchy { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Add a new display item
        /// </summary>
        /// <param name="machineName">name of the sender machine</param>
        /// <param name="appDomain">name of the sender appdomain</param>
        /// <param name="time">sender time</param>
        /// <param name="text">display text</param>
        /// <param name="parentID">parent display item id</param>
        /// <returns>the new created id for the item</returns>
        public string AddNew(string machineName, string appDomain, string time, string text, string parentID = null)
        {
            _idCounter++;
            ConsoleViewItem newItem = new ConsoleViewItem(_idCounter.ToString(), machineName, appDomain, time, text);
            Plain.Add(newItem);
            if (!String.IsNullOrWhiteSpace(parentID))
            {
                ConsoleViewItem parentItem = LookupForItem(parentID);
                if (null != parentItem)
                    parentItem.Items.Add(newItem);
                else
                    Hierarchy.Add(newItem);
            }
            else
                Hierarchy.Add(newItem);

            return _idCounter.ToString();
        }

        /// <summary>
        /// Clear display item content
        /// </summary>
        public void Clear()
        {
            _idCounter = 0;
            Plain.Clear();
            Hierarchy.Clear();
        }

        /// <summary>
        /// Create a System.String to display the items
        /// </summary>
        /// <param name="style">view/format options</param>
        /// <param name="showTime">show sender time</param>
        /// <param name="showMachine">show sender machine</param>
        /// <param name="showAppDomain">show sender appDomain</param>
        /// <returns>String to display</returns>
        public string CreateText(ConsoleViewStyle style, bool showTime, bool showMachine, bool showAppDomain)
        {
            switch (style)
            {
                case ConsoleViewStyle.Plain:
                    return CreatePlainText(showTime, showMachine, showAppDomain);
                case ConsoleViewStyle.PlainReverse:
                    return CreatePlainReverseText(showTime, showMachine, showAppDomain);
                case ConsoleViewStyle.Hierarchy:
                    return CreateHierarchyText(showTime, showMachine, showAppDomain);
                default:
                    throw new ArgumentOutOfRangeException("style");
            }
        }

        /// <summary>
        /// Creates a System.String to display the items
        /// </summary>
        /// <param name="showTime">show sender time</param>
        /// <param name="showMachine">show sender machine</param>
        /// <param name="showAppDomain">show sender appDomain</param>
        /// <returns>String to display</returns>
        public string CreatePlainText(bool showTime, bool showMachine, bool showAppDomain)
        {
            StringBuilder sb = new StringBuilder(_stringBuilderDefaultSize);
            foreach (var item in Plain)
            {
                string timeString = "[" + item.Time + "]";
                string machineString = "[" + (String.IsNullOrWhiteSpace(item.Machine) ? _emptyMachineDefault : item.Machine) + "]";
                string domainString = "[" + (String.IsNullOrWhiteSpace(item.AppDomain) ? _emptyAppDomainDefault : item.AppDomain) + "]";

                string line = (true == showTime ? timeString : "") +
                    (true == showMachine ? machineString : "") +
                    (true == showAppDomain ? domainString : "") +
                    item.Text;

                sb.Append(line + Environment.NewLine);
            }
            return sb.ToString();
        }

        /// <summary>
        ///  Creates a System.String to display the items in reverse order
        /// </summary>
        /// <param name="showTime">show sender time</param>
        /// <param name="showMachine">show sender machine</param>
        /// <param name="showAppDomain">show sender appDomain</param>
        /// <returns>String to display</returns>
        public string CreatePlainReverseText(bool showTime, bool showMachine, bool showAppDomain)
        {
            StringBuilder sb = new StringBuilder(_stringBuilderDefaultSize);
            for (int i = Plain.Count; i > 0; i--)
            {
                ConsoleViewItem item = Plain[i-1];

                string timeString = "[" + item.Time + "]";
                string machineString = "[" + (String.IsNullOrWhiteSpace(item.Machine) ? _emptyMachineDefault : item.Machine) + "]";
                string domainString = "[" + (String.IsNullOrWhiteSpace(item.AppDomain) ? _emptyAppDomainDefault : item.AppDomain) + "]";

                string line = (true == showTime ? timeString : "") +
                    (true == showMachine ? machineString : "") +
                    (true == showAppDomain ? domainString : "") +
                    item.Text;

                sb.Append(line + Environment.NewLine);
            }
            return sb.ToString();
        }

        /// <summary>
        ///  Creates a System.String to display the items in a hierarchy order 
        /// </summary>
        /// <param name="showTime">show sender time</param>
        /// <param name="showMachine">show sender machine</param>
        /// <param name="showAppDomain">show sender appDomain</param>
        /// <returns>String to display</returns>
        public string CreateHierarchyText(bool showTime, bool showMachine, bool showAppDomain)
        {
            StringBuilder sb = new StringBuilder(_stringBuilderDefaultSize);
            int numSpaces = 0;
            foreach (var item in Hierarchy)
            {
                string timeString = "[" + item.Time + "]";
                string machineString = "[" + (String.IsNullOrWhiteSpace(item.Machine) ? _emptyMachineDefault : item.Machine) + "]";
                string domainString = "[" + (String.IsNullOrWhiteSpace(item.AppDomain) ? _emptyAppDomainDefault : item.AppDomain) + "]";

                string line = (true == showTime ? timeString : "") +
                    (true == showMachine ? machineString : "") +
                    (true == showAppDomain ? domainString : "") +
                    item.Text + Environment.NewLine;

                sb.Append(line);

                string text = CreateHierarchyText(item, numSpaces, showTime, showMachine, showAppDomain);
                if (!String.IsNullOrWhiteSpace(text))
                    sb.Append(text);
            }
            return sb.ToString();
        }

        #endregion

        #region Private Methods

        private ConsoleViewItem LookupForItem(string id)
        {
            foreach (var item in Hierarchy)
            {
                if (item.ID == id)
                    return item;
                ConsoleViewItem child = LookupForItem(item, id);
                if (null != child)
                    return child;
            }

            return null;
        }
     
        private ConsoleViewItem LookupForItem(ConsoleViewItem element, string id)
        {
            foreach (var item in element.Items)
            {
                if (item.ID == id)
                    return item;
                ConsoleViewItem child = LookupForItem(item, id);
                if (null != child)
                    return child;
            }

            return null;
        }

        private string CreateHierarchyText(ConsoleViewItem element, int numSpaces, bool showTime, bool showMachine, bool showAppDomain)
        {
            string lines = String.Empty;
            numSpaces++;
            foreach (var item in element.Items)
            {
                string timeString = "[" + item.Time + "]";
                string machineString = "[" + (String.IsNullOrWhiteSpace(item.Machine) ? _emptyMachineDefault : item.Machine) + "]";
                string domainString = "[" + (String.IsNullOrWhiteSpace(item.AppDomain) ? _emptyAppDomainDefault : item.AppDomain) + "]";

                string line = (true == showTime ? timeString : "") +
                    (true == showMachine ? machineString : "") +
                    (true == showAppDomain ? domainString : "") +
                    CreateSpaceString(numSpaces) + item.Text + Environment.NewLine;

                line += CreateHierarchyText(item, numSpaces+1, showTime, showMachine, showAppDomain);
                lines += line;

            }
            return lines;
        }

        private string CreateSpaceString(int numSpaces)
        {
            string result = "";
            for (int i = 0; i < numSpaces; i++)
                result += _spaceString;
            return result;
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("{0}Plain Items, {1} Hierachy Items", Plain.Count, Hierarchy.Count);
        }

        #endregion
    }
}
