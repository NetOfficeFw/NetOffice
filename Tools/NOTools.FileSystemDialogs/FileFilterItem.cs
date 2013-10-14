using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    internal class FileFilterItem
    {
        internal FileFilterItem(string name, string filter)
        {
            Name = name;
            Filter = filter;
        }

        public string Name { get; private set; }
        public string Filter { get; private set; }

        public static FileFilterItem[] CreateFromFilterString(string filterString)
        {
            if (String.IsNullOrWhiteSpace(filterString))
                return new FileFilterItem[1] { new FileFilterItem("", "*.*") };

            string[] array = filterString.Split( new string[]{"|"}, StringSplitOptions.None);
            if (array.Length % 2 != 0)
                throw new FormatException("File filter must be like - txt files (*.txt)|*.txt|All files (*.*)|*.*");

            List<FileFilterItem> list = new List<FileFilterItem>();
            for (int i = 0; i < array.Length; i = i + 2)
                list.Add( new FileFilterItem(array[i], array[i+1]));

            return list.ToArray();
        }
    }
}
