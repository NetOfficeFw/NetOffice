using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using NetOffice.DeveloperToolbox.Utils.Registry;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    public class RegistrySearch
    {
        public RegistrySearch(IEnumerable<UtilsRegistry> rootKeys, UtilsRegistryKey startFromKey, string expression)
        {
            if (null == rootKeys)
                throw new ArgumentNullException("rootKeys");
            if (null == startFromKey)
                throw new ArgumentNullException("startFromKey");
            if (String.IsNullOrWhiteSpace(expression))
                throw new ArgumentNullException("expression");
            Expression = expression;
            RootKeys = rootKeys;
            StartFromKey = startFromKey;
            StartFromHive = startFromKey.Root;
            StartFromKeyIsMatchExpression = ValidateStartFormKeyIsMatchingExpression(startFromKey);
        }

        public IEnumerable<UtilsRegistry> RootKeys { get; private set; }

        public UtilsRegistry StartFromHive { get; private set; }

        public UtilsRegistryKey StartFromKey { get; private set; }

        public string Expression { get; private set; }

        public bool StartFromKeyIsMatchExpression { get; private set; }

        public bool CurrentlySearching { get; private set; }

        public bool SearchPassed { get; private set; }

        public UtilsRegistryKey Result { get; private set; }

        public IEnumerable<UtilsRegistryEntry> ResultEntries { get; private set; }

        public bool Search()
        {
            try
            {
                Result = null;
                ResultEntries = null;
                CurrentlySearching = true;

                List<UtilsRegistryEntry> resultEntries = new List<UtilsRegistryEntry>();
                UtilsRegistryKey startKey = null;

                if (StartFromKeyIsMatchExpression)
                    startKey = StartFromKey.Next();
                else
                    startKey = startKey = StartFromKey;

                string padding = "";
                bool found = false;

                while (null != startKey)
                {
                    var entry = startKey;
                    found = SearchInternal(entry, ref resultEntries, padding);
                    if (found)
                        break;

                    startKey = startKey.Next();
                }

                if (!found)
                {
                    padding = String.Empty;
                    var rootKeys = GetRootKeysBelowStartKey();
                    foreach (var rootKey in rootKeys)
                    {
                        startKey = rootKey.Key;
                        while (null != startKey)
                        {
                            var entry = startKey;
                            found = SearchInternal(entry, ref resultEntries, padding);
                            if (found)
                                break;

                            startKey = startKey.Next();
                        }
                    }
                }

                if (found)
                    ResultEntries = resultEntries;

                return found;
            }
            catch
            {
                throw;
            }
            finally
            {
                CurrentlySearching = false;
                SearchPassed = true;
            }
        }

        private bool ValidateStartFormKeyIsMatchingExpression(UtilsRegistryKey key)
        {
            bool found = key.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1;
            if (found)
                return true;

            var entries = key.Entries;
            foreach (UtilsRegistryEntry item in entries)
            {
                string valueString = null != item.Value ? item.Value.ToString() : String.Empty;
                if ((item.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1) ||
                    (valueString.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1))
                {
                    return true;
                }
            }
            return false;
        }

        private bool SearchInternal(UtilsRegistryKey key, ref List<UtilsRegistryEntry> resultEntries, string padding)
        {
            //Console.WriteLine("{0}{1}", padding, key.Name);

            bool found = key.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1;
            if (found)
                Result = key;
            var entries = key.Entries;

            foreach (UtilsRegistryEntry item in entries)
            {
                string valueString = null != item.Value ? item.Value.ToString() : String.Empty;
                if ((item.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1) ||
                    (valueString.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1))
                {
                    Result = key;
                    resultEntries.Add(item);
                    found = true;
                }
            }
            return found;
        }

        private IEnumerable<UtilsRegistry> GetRootKeysBelowStartKey()
        {
            bool sameRootPassed = false;
            List<UtilsRegistry> result = new List<UtilsRegistry>();
            foreach (var item in RootKeys)
            {
                if (sameRootPassed)
                    result.Add(item);
                else
                    sameRootPassed = StartFromHive.Name == item.Name;
            }
            return result;
        }
    }
}