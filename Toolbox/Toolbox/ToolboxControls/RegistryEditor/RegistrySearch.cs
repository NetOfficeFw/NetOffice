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
        public RegistrySearch(IEnumerable<UtilsRegistry> rootKeys, RegistryKey startFromHive, string startFromPath, bool startFromPathIsTopLevel, string expression, bool startFromNextPossiblePath)
        {
            if (null == rootKeys)
                throw new ArgumentNullException("rootKeys");
            if (null == startFromPath)
                throw new ArgumentNullException("startFromPath");
            if (String.IsNullOrWhiteSpace(expression))
                throw new ArgumentNullException("expression");
            RootKeys = rootKeys;
            string path = startFromPath.Substring(0, startFromPath.LastIndexOf("\\"));
            string name = startFromPath.Substring(startFromPath.LastIndexOf("\\")+ "\\".Length);
            StartFrom = new UtilsRegistry(startFromHive, startFromPath);
            StartFromParent = new UtilsRegistry(startFromHive, path);
            StartFromName = name;
            Expression = expression;
            StartFromNextPossiblePath = startFromNextPossiblePath;
            StartFromPathIsTopLevel = startFromPathIsTopLevel;
        }

        public IEnumerable<UtilsRegistry> RootKeys { get; private set; }

        public UtilsRegistry StartFrom { get; private set; }

        public UtilsRegistry StartFromParent { get; private set; }

        public string StartFromName { get; private set; }

        public string Expression { get; private set; }

        public bool StartFromNextPossiblePath { get; private set; }

        public bool StartFromPathIsTopLevel { get; private set; }

        public bool CurrentlySearching { get; private set; }

        public bool SearchPassed { get; private set; }

        public UtilsRegistry Result { get; private set; }

        public IEnumerable<UtilsRegistryEntry> ResultEntries { get; private set; }

        public bool Search()
        {
            try
            {
                Result = null;
                ResultEntries = null;
                CurrentlySearching = true;

                List<UtilsRegistryEntry> resultEntries = new List<UtilsRegistryEntry>();
                UtilsRegistryKey[] keys = null;
                int keysStartIndex = 0;

                if (StartFromPathIsTopLevel)
                {
                    keys = StartFrom.Key.Keys.ToArray();
                    keysStartIndex = 0;
                }
                else
                {
                    keys = StartFromParent.Key.Keys.ToArray();
                    keysStartIndex = GetStartFromEntriesStartIndex(keys);
                }

                bool found = false;
                for (int i = keysStartIndex; i < keys.Length; i++)
                {
                    var entry = keys[i];
                    found = SearchInternal(entry, ref resultEntries);
                    if (!found)
                    {
                        var rootKeys = GetRootKeysBelowStartKey();
                        foreach (var item in rootKeys)
                        {
                            found = SearchInternal(item.Key, ref resultEntries);
                            if (found)
                                break;
                        }
                    }
                    else
                        break;
                }
                if(found)
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

        private bool SearchInternal(UtilsRegistryKey key, ref List<UtilsRegistryEntry> resultEntries)
        {
            bool found = key.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1;
            if (found)
                Result = new UtilsRegistry(key.Root.HiveKey, key.Path);
            var entries = key.Entries;

            foreach (UtilsRegistryEntry item in entries)
            {
                string valueString = null != item.Value ? item.Value.ToString() : String.Empty;
                if ((item.Name.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1) ||
                    (valueString.IndexOf(Expression, StringComparison.InvariantCultureIgnoreCase) > -1))
                {
                    resultEntries.Add(item);
                    Result = new UtilsRegistry(key.Root.HiveKey, key.Path);
                    found = true;
                }
            }

            if (!found)
            {
                foreach (var item in key.Keys)
                {
                    if (SearchInternal(item, ref resultEntries))
                        return true;
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
                    sameRootPassed = StartFromParent.HiveKey == item.HiveKey;
            }
            return result;
        }

        private int GetStartFromEntriesStartIndex(UtilsRegistryKey[] keys)
        {
            int i = 0;
            foreach (var key in keys)
            {
                if (key.Name == StartFromName)
                    break;
                i++;
            }

            if (StartFromNextPossiblePath)
                i++;
            return i;
        }
    }
}