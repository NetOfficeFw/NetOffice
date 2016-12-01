using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu item type
    /// </summary>
    public enum TrayMenuItemType
    {
        /// <summary>
        /// Standard Item
        /// </summary>
        Item = 0,

        /// <summary>
        /// Standard Label
        /// </summary>
        Label = 1,

        /// <summary>
        /// Link Label
        /// </summary>
        LinkLabel = 2,

        /// <summary>
        /// Standard Button
        /// </summary>
        Button = 3,
       
        /// <summary>
        /// Standard Text Box
        /// </summary>
        TextBox = 4,

        /// <summary>
        /// Checkable Item 
        /// </summary>
        CheckBox = 5,

        /// <summary>
        /// Progress Bar
        /// </summary>
        Progress = 6,

        /// <summary>
        /// Drop Down List
        /// </summary>
        DropDownList = 7,

        /// <summary>
        /// Separator Line
        /// </summary>
        Separator = 8,

        /// <summary>
        /// Custom Element
        /// </summary>
        Custom = 9,
        
        /// <summary>
        /// Diagnostics Monitor
        /// </summary>
        Monitor = 20,

        /// <summary>
        /// Tray Menu Auto Close Check Item
        /// </summary>
        AutoClose = 21,

        /// <summary>
        /// Tray Menu Close Item
        /// </summary>
        Close = 22
    }

    /// <summary>
    /// TrayMenuItem and its dervied classes use this attribute to identifier
    /// the class item type at runtime
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ItemTypeAttribute : System.Attribute
    {
        /// <summary>
        /// Runtime supported item type
        /// </summary>
        public readonly TrayMenuItemType Type;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">runtime supported item type</param>
        public ItemTypeAttribute(TrayMenuItemType type)
        {
            Type = type;
        }
    }
}
