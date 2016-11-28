using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu close item to close the tray menu
    /// </summary>
    [ItemType(TrayMenuItemType.Close)]
    public class TrayMenuCloseItem : TrayMenuItem
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuCloseItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.Close;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuCloseItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.Close;
        }

        #endregion
    }
}
