using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Runtime;
using System.Collections;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu item
    /// </summary>
    [ItemType(TrayMenuItemType.Item)]
    public class TrayMenuItem
    {
        #region Fields

        private TrayMenu _owner;

        private string _text;

        private bool _visible;

        private string _toolTipText;

        private Image _image;

        private Color _backColor = Color.FromKnownColor(KnownColor.Control);

        private Color _foreColor = Color.FromKnownColor(KnownColor.ControlText);

        private Font _font;

        private bool _enabled = true;

        private ContentAlignment _textAlign;

        private ContentAlignment _imageAlign;

        private Padding _padding;

        private TrayMenuItems _items;

        private object _itemsLock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuItem(TrayMenu owner, string text)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
            Text = text;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuItem(TrayMenu owner, string text, bool visible)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
            Text = text;
            Visible = visible;
        }

        #endregion

        #region Properties
         
        /// <summary>
        /// Owner Menu
        /// </summary>
        protected internal TrayMenu Owner
        {
            get
            {
                return _owner;
            }
        }
         
        /// <summary>
        /// Optional Child Items
        /// </summary>
        public virtual TrayMenuItems Items
        {
            get
            {
                lock (_itemsLock)
                {
                    if (null == _items)
                        _items = OnCreateMenuItems();
                }
                return _items;
            }
        }

        /// <summary>
        /// Background color
        /// </summary>
        public virtual Color BackColor
        {
            get
            {
                return _backColor;
            }
            set
            {
                if (value != _backColor)
                {
                    _backColor = value;
                    _owner.OnItemBackColorChanged(this);
                }
            }
        }

        /// <summary>
        /// Fore/Font color
        /// </summary>
        public virtual Color ForeColor
        {
            get
            {
                return _foreColor;
            }
            set
            {
                if (value != _foreColor)
                {
                    _foreColor = value;
                    _owner.OnItemForeColorChanged(this);
                }
            }
        }

        /// <summary>
        /// Item Font
        /// </summary>
        public virtual Font Font
        {
            get
            {
                return _font;
            }
            set
            {
                if (value != _font)
                {
                    _font = value;
                    _owner.OnItemFontChanged(this);
                }
            }
        }
       
        /// <summary>
        /// Get or set item visibility
        /// </summary>
        public virtual bool Visible
        {
            get
            {
                return _visible;
            }
            set
            {
                if (value != _visible)
                {
                    _visible = value;
                    _owner.OnItemVisibleChanged(this);
                }
            }
        }
        
        /// <summary>
        /// Item Enabled State
        /// </summary>
        public virtual bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {
                if (value != _enabled)
                {
                    _enabled = value;
                    _owner.OnItemEnabledChanged(this);
                }
            }
        }

        /// <summary>
        /// Shown caption
        /// </summary>
        public virtual string Text
        {
            get
            {
                return _text;
            }
            set
            {
                if (value != _text)
                {
                    _text = value;
                    _owner.OnItemTextChanged(this);
                }
            }
        }

        /// <summary>
        /// Shown Text Alignment
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public virtual ContentAlignment TextAlign
        {
            get
            {
                return _textAlign;
            }
            set
            {
                if (value != _textAlign)
                {
                    _textAlign = value;
                    _owner.OnItemTextAlignChanged(this);
                }
            }
        }

        /// <summary>
        /// Shown Tooltip
        /// </summary>
        public virtual string ToolTipText
        {
            get
            {
                return _toolTipText;
            }
            set
            {
                if (value != _toolTipText)
                {
                    _toolTipText = value;
                    _owner.OnItemToolTipTextChanged(this);
                }
            }
        }


        /// <summary>
        /// Shown Image
        /// </summary>
        public virtual Image Image
        {
            get
            {
                return _image;
            }
            set
            {
                if (value != _image)
                {
                    _image = value;
                    _owner.OnItemImageChanged(this);
                }
            }
        }

        /// <summary>
        /// Shown Image Alignment
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public virtual ContentAlignment ImageAlign
        {
            get
            {
                return _imageAlign;
            }
            set
            {
                if (value != _imageAlign)
                {
                    _imageAlign = value;
                    _owner.OnItemImageAlignChanged(this);
                }
            }
        }

        /// <summary>
        /// Padding Space
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public virtual Padding Padding
        {
            get
            {
                return _padding;
            }
            set
            {
                if (value != _padding)
                {
                    _padding = value;
                    _owner.OnItemPaddingChanged(this);
                }
            }
        }

        /// <summary>
        /// Instance Category Type
        /// </summary>
        public TrayMenuItemType ItemType { get; protected internal set; }

        /// <summary>
        /// Well known any tag
        /// </summary>
        public object Tag { get; set; }

        /// <summary>
        /// Same as Tag
        /// </summary>
        public object Extender { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Initialize values from ui element
        /// </summary>
        /// <param name="font">element font</param>
        /// <param name="textAlign">text align</param>
        /// <param name="imageAlign">image align</param>
        /// <param name="padding">padding space</param>
        internal void SetupElements(Font font, ContentAlignment textAlign, ContentAlignment imageAlign, Padding padding)
        {
            _font = font;
            _textAlign = textAlign;
            _imageAlign = imageAlign;
            _padding = padding;
        }

        /// <summary>
        /// Creates an instance of TrayMenuItems
        /// </summary>
        /// <returns>TrayMenuItems instance</returns>
        protected internal virtual TrayMenuItems OnCreateMenuItems()
        {
            return new TrayMenuItems(Owner, this);
        }

        #endregion
    }
}
