using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a dumy collection of tray menu items
    /// </summary>
    public class TrayMenuStubItems : TrayMenuItems
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">collection owner</param>
        internal TrayMenuStubItems(TrayMenu owner) : base(owner)
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">collection owner</param>
        /// <param name="parent">parent item instance</param>
        internal TrayMenuStubItems(TrayMenu owner, TrayMenuItem parent) : base(owner, parent)
        {

        }

        #endregion

        #region Overrides

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public override IEnumerable<TrayMenuItem> Add(params string[] text)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        ///  Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text, bool visible)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <param name="image"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text, bool visible, Image image)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <param name="image"></param>
        /// <param name="itemType"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text, bool visible, Image image, TrayMenuItemType itemType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <param name="itemType"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text, bool visible, TrayMenuItemType itemType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="text"></param>
        /// <param name="itemType"></param>
        /// <returns></returns>
        public override TrayMenuItem Add(string text, TrayMenuItemType itemType)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="text"></param>
        /// <returns></returns>
        public override IEnumerable<T> Add<T>(params string[] text)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="text"></param>
        /// <returns></returns>
        public override T Add<T>(string text)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <returns></returns>
        public override T Add<T>(string text, bool visible)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="text"></param>
        /// <param name="visible"></param>
        /// <param name="image"></param>
        /// <returns></returns>
        public override T Add<T>(string text, bool visible, Image image)
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        public override void Clear()
        {
            ;
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public override int IndexOf(TrayMenuItem item)
        {
            return -1;
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public override TrayMenuItem this[int index]
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public override bool Remove(TrayMenuItem item)
        {
            return false;
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        public override int Count
        {
            get
            {
                return 0;
            }
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <param name="index"></param>
        public override void RemoveAt(int index)
        {
            ;
        }

        /// <summary>
        /// Not Supported
        /// </summary>
        /// <returns></returns>
        public override IEnumerator<TrayMenuItem> GetEnumerator()
        {
            return new TrayMenuItem[0].GetEnumerator() as IEnumerator<TrayMenuItem>;
        }

        #endregion
    }
}
