using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu progressbar item
    /// </summary>
    [ItemType(TrayMenuItemType.Progress)]
    public class TrayMenuProgressItem : TrayMenuItem
    {
        #region Nested

        /// <summary>
        ///  Specifies the style that a ProgressBar uses to indicate the progress of an operation.
        /// </summary>
        public enum ProgressBarStyle
        {
            /// <summary>
            /// Indicates progress by increasing the number of segmented blocks in a ProgressBar.
            /// </summary>
            Blocks = 0,

            /// <summary>
            /// Indicates progress by increasing the size of a smooth, continuous bar in a ProgressBar.
            /// </summary>
            Continuous = 1,

            /// <summary>
            /// Indicates progress by continuously scrolling a block across a ProgressBar in a marquee fashion.
            /// </summary>
            Marquee = 2
        }

        #endregion

        #region Fields

        private int _minimum;
        private int _maximum;
        private int _value;
        private ProgressBarStyle _style;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuProgressItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.Progress;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuProgressItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.Progress;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Minimum Allowed Value
        /// </summary>
        public int Minimum
        {
            get
            {
                return _minimum;
            }
            set
            {
                if (value != _minimum)
                {
                    _minimum = value;
                    _minimum = Owner.OnProgressItemMinimumChanged(this);
                }
            }
        }

        /// <summary>
        /// Maximum Allowed Value
        /// </summary>
        public int Maximum
        {
            get
            {
                return _maximum;
            }
            set
            {
                if (value != _maximum)
                {
                    _maximum = value;
                    _maximum = Owner.OnProgressItemMaximumChanged(this);
                }
            }
        }

        /// <summary>
        /// Current shown value
        /// </summary>
        public int Value
        {
            get
            {
                return _value;
            }
            set
            {
                if (value != _value)
                {
                    _value = value;
                    _value = Owner.OnProgressItemValueChanged(this);
                }
            }
        }

        /// <summary>
        /// ProgressBar Style
        /// </summary>
        public ProgressBarStyle Style
        {
            get
            {
                return _style;
            }
            set
            {
                if (value != _style)
                {
                    _style = value; 
                    Owner.OnProgressItemStyleChanged(this);
                }
            }
        }

        #endregion

        #region Methods

        internal void SetProgressElements(int minimum, int maximum, int value, ProgressBarStyle style)
        {
            _minimum = minimum;
            _maximum = maximum;
            _value = value;
            _style = style;
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Optional child items which is not supported in this item type
        /// </summary>
        [System.ComponentModel.Browsable(false), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public override TrayMenuItems Items
        {
            get
            {
                return base.Items;
            }
        }
        /// <summary>
        /// Creates a new items collection
        /// </summary>
        /// <returns>collection instance</returns>
        protected internal override TrayMenuItems OnCreateMenuItems()
        {
            return new TrayMenuStubItems(Owner, this);
        }

        #endregion
    }
}
