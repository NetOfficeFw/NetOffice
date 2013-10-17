using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Represents an available filter in the file type combo box
    /// </summary>
    internal class FileFilterItem
    {
        internal FileFilterItem(string name, string filter)
        {
            Name = name;
            Filter = filter;
        }

        /// <summary>
        /// The first argument
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// The second argument
        /// </summary>
        public string Filter { get; private set; }

         /// <summary>
         /// Creates a filterfilter array from string
         /// </summary>
         /// <param name="filterString">given string as any</param>
         /// <returns>new created array</returns>
        internal static FileFilterItem[] CreateFromFilterString(string filterString)
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
