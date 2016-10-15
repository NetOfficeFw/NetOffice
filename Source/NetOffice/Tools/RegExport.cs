using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace NetOffice.Tools
{
    /// <summary>
    /// Registry Export Definition
    /// </summary>
    [Guid("DE1590FF-EA17-4A2C-AE2B-62AF3EFD887F")]
    public class RegExport : Dictionary<string, IList<RegExportValue>>
    {
        /// <summary>
        ///  Add a new key to the instancc
        /// </summary>
        /// <param name="key">target unique key</param>
        /// <returns>value list</returns>
        public IList<RegExportValue> Add(string key)
        {
            var list = new List<RegExportValue>();
            base.Add(key, list);
            return list;
        }

        /// <summary>
        /// Add a new key to the instancc
        /// </summary>
        /// <param name="key">target unique key</param>
        /// <param name="values">target values</param>
        /// <returns>value list</returns>
        public IList<RegExportValue> Add(string key, IEnumerable<RegExportValue> values)
        {
            var list = new List<RegExportValue>();
            foreach (RegExportValue item in values)
                list.Add(item);
            base.Add(key, list);           
            return list;
        }
    }

    /// <summary>
    /// Represents a registry value
    /// </summary>
    [Guid("B469F46B-385B-4BE3-9420-6090CC4D145D")]
    public class RegExportValue
    {
        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        public RegExportValue()
        {
        }
         
        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="kind">target value kind</param>
        /// <param name="value">target value</param>
        public RegExportValue(string name, RegistryValueKind kind, object value)
        {
            Name = name;
            Kind = kind;
            Value = value;
        }

        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        /// <param name="value">target value</param>
        public RegExportValue(object value)
        {
            Kind = RegistryValueKind.String;
            Value = value;
        }
        
        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="value">target value</param>
        public RegExportValue(string name, object value)
        {
            Name = name;
            Kind = RegistryValueKind.String;
            Value = value;
        }

        /// <summary>
        /// Value name, can be null for(default) value
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Value Kind
        /// </summary>
        public RegistryValueKind Kind { get; set; }

        /// <summary>
        /// Value
        /// </summary>
        public object Value { get; set; }
    }
}
