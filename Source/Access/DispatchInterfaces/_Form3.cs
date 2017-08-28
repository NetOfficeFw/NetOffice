using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Form3 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class _Form3 : _Form2 , IEnumerable<object>
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_Form3);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Form3(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form3(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FormName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string RecordSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FilterOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OrderBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OrderBy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool OrderByOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OrderByOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderByOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowFilters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFilters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFilters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DefaultView
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DefaultView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte ViewsAllowed
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ViewsAllowed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewsAllowed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool AllowEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 DefaultEditing
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DefaultEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowEdits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowDeletions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AllowAdditions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DataEntry
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataEntry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte AllowUpdating
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "AllowUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte RecordsetType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "RecordsetType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordsetType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte RecordLocks
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "RecordLocks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordLocks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte ScrollBars
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ScrollBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool RecordSelectors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RecordSelectors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSelectors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool NavigationButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NavigationButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NavigationButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DividingLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DividingLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DividingLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AutoResize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AutoCenter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoCenter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoCenter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PopUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PopUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PopUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Modal
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Modal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Modal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte BorderStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ControlBox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ControlBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MinButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MinButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MaxButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaxButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte MinMaxButtons
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "MinMaxButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinMaxButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool CloseButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CloseButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CloseButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool WhatsThisButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WhatsThisButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WhatsThisButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Width
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Picture
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Picture");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Picture", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureSizeMode
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureSizeMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureAlignment
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureAlignment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PictureTiling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte Cycle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Cycle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Cycle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string MenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Toolbar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Toolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Toolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ShortcutMenu
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShortcutMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortcutMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ShortcutMenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridX
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridY
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool LayoutForPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutForPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutForPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FastLaserPrinting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FastLaserPrinting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FastLaserPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string HelpFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 HelpContextId
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 RowHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string DatasheetFontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DatasheetFontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DatasheetFontHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DatasheetFontHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DatasheetFontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DatasheetFontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DatasheetFontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DatasheetFontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool DatasheetFontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DatasheetFontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetFontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte TabularCharSet
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TabularCharSet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabularCharSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DatasheetGridlinesBehavior
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetGridlinesBehavior");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetGridlinesBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetGridlinesColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetGridlinesColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetGridlinesColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DatasheetCellsEffect
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetCellsEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetCellsEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ShowGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 DatasheetBackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetBackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Hwnd
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hwnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hwnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Page
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Page");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Pages
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Pages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Pages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ZoomControl
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ZoomControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZoomControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Painting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Painting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Painting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtMip
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtMip");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtMip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevMode
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevMode");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevNames
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevNames");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FrozenColumns
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FrozenColumns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FrozenColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Bookmark
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Bookmark");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte TabularFamily
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TabularFamily");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabularFamily", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string PaletteSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PaletteSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PaletteSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PaintPalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PaintPalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PaintPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMenu
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object OpenArgs
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OpenArgs");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OpenArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ConnectSynch
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ConnectSynch");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectSynch", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnCurrent
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnCurrent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnCurrent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterInsert
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterInsert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterInsert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeUpdate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterUpdate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDirty
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDelete
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDelete");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDelete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BeforeDelConfirm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeDelConfirm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string AfterDelConfirm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterDelConfirm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterDelConfirm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnOpen
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnLoad
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLoad");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnResize
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnUnload
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUnload");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUnload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnClose
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClose");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnActivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnActivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDeactivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDeactivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnGotFocus
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnGotFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnGotFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnLostFocus
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLostFocus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLostFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDblClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDblClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDblClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseDown
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseMove
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseMove");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseMove", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnMouseUp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyDown
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyUp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnKeyPress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyPress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyPress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool KeyPreview
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KeyPreview");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeyPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnError
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnApplyFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnApplyFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnApplyFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnTimer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnTimer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnTimer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 TimerInterval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TimerInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TimerInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Dirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 WindowWidth
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 WindowHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WindowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentView
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentSectionTop
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentSectionTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentSectionTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 CurrentSectionLeft
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrentSectionLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentSectionLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 SelHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PictureData
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PictureData");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PictureData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 InsideHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InsideHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InsideHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 InsideWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InsideWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InsideWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PicturePalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PicturePalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PicturePalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool HasModule
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasModule");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasModule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 acHiddenCurrentPage
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "acHiddenCurrentPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "acHiddenCurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte Orientation
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool AllowDesignChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDesignChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDesignChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ServerFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ServerFilterByForm
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerFilterByForm");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilterByForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 MaxRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string UniqueTable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueTable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UniqueTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ResyncCommand
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResyncCommand");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResyncCommand", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string InputParameters
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InputParameters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InputParameters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool MaxRecButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaxRecButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 NewRecord
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "NewRecord");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control ActiveControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ActiveControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Control get_DefaultControl(Int32 controlType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "DefaultControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultControl
		/// </summary>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_DefaultControl")]
		public NetOffice.AccessApi.Control DefaultControl(Int32 controlType)
		{
			return get_DefaultControl(controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Dynaset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Dynaset");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object RecordsetClone
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "RecordsetClone");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Recordset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Recordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 12)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Section get__SectionOld(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "_SectionOld", NetOffice.AccessApi.Section.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 12
		/// Alias for get__SectionOld
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 12), Redirect("get__SectionOld")]
		public NetOffice.AccessApi.Section _SectionOld(object index)
		{
			return get__SectionOld(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Form Form
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Form>(this, "Form", NetOffice.AccessApi.Form.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Module Module
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Module>(this, "Module", NetOffice.AccessApi.Module.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", NetOffice.AccessApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Control ConnectControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ConnectControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Controls Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Controls>(this, "Controls", NetOffice.AccessApi.Controls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 SubdatasheetHeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "SubdatasheetHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubdatasheetHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool SubdatasheetExpanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SubdatasheetExpanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubdatasheetExpanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte DatasheetBorderLineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetBorderLineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetBorderLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte DatasheetColumnHeaderUnderlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetColumnHeaderUnderlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetColumnHeaderUnderlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte HorizontalDatasheetGridlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "HorizontalDatasheetGridlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HorizontalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte VerticalDatasheetGridlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "VerticalDatasheetGridlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VerticalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowTop
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowTop");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowLeft
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnUndo
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUndo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUndo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnRecordExit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnRecordExit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnRecordExit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		public object PivotTable
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PivotTable");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		public object ChartSpace
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ChartSpace");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi._Printer Printer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi._Printer>(this, "Printer", NetOffice.AccessApi._Printer.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Printer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool Moveable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Moveable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Moveable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeginBatchEdit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeginBatchEdit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeginBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string UndoBatchEdit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UndoBatchEdit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UndoBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeBeginTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeBeginTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterBeginTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterBeginTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeCommitTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeCommitTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterCommitTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterCommitTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string RollbackTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RollbackTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RollbackTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowFormView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFormView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFormView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowDatasheetView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDatasheetView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDatasheetView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowPivotTableView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPivotTableView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPivotTableView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowPivotChartView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPivotChartView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPivotChartView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnConnect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnConnect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnDisconnect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDisconnect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDisconnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string PivotTableChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PivotTableChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PivotTableChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string Query
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Query");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Query", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeQuery
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeQuery");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeQuery", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string SelectionChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SelectionChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelectionChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandBeforeExecute
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandBeforeExecute");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandBeforeExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandChecked
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandEnabled
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandExecute
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandExecute");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DataSetChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSetChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSetChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeScreenTip
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeScreenTip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeScreenTip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterFinalRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterFinalRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterFinalRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterLayout
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string MouseWheel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MouseWheel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MouseWheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string ViewChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ViewChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DataChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool FetchDefaults
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FetchDefaults");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FetchDefaults", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool BatchUpdates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BatchUpdates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BatchUpdates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte CommitOnClose
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "CommitOnClose");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommitOnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CommitOnNavigation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CommitOnNavigation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommitOnNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool UseDefaultPrinter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseDefaultPrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseDefaultPrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string RecordSourceQualifier
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSourceQualifier");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSourceQualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool FilterOnLoad
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterOnLoad");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterOnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool OrderByOnLoad
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OrderByOnLoad");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderByOnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcSplitFormOrientation SplitFormOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormOrientation>(this, "SplitFormOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitFormOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcSplitFormDatasheet SplitFormDatasheet
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormDatasheet>(this, "SplitFormDatasheet");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitFormDatasheet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool SplitFormSplitterBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SplitFormSplitterBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SplitFormSplitterBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcSplitFormPrinting SplitFormPrinting
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormPrinting>(this, "SplitFormPrinting");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitFormPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool SplitFormSplitterBarSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SplitFormSplitterBarSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SplitFormSplitterBarSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public string NavigationCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NavigationCaption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NavigationCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnCurrentMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnCurrentMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnCurrentMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeInsertMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeInsertMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeInsertMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterInsertMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterInsertMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterInsertMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeUpdateMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeUpdateMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeUpdateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterUpdateMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterUpdateMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterUpdateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDirtyMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDirtyMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDirtyMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDeleteMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDeleteMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDeleteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeDelConfirmMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeDelConfirmMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeDelConfirmMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterDelConfirmMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterDelConfirmMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterDelConfirmMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnOpenMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnOpenMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnOpenMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnLoadMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLoadMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLoadMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnResizeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnResizeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnResizeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnUnloadMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUnloadMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUnloadMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnCloseMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnCloseMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnCloseMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnActivateMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnActivateMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnActivateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDeactivateMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDeactivateMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDeactivateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnGotFocusMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnGotFocusMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnGotFocusMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnLostFocusMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnLostFocusMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnLostFocusMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnClickMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClickMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClickMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDblClickMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDblClickMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDblClickMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMouseDownMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseDownMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseDownMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMouseMoveMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseMoveMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseMoveMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMouseUpMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMouseUpMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMouseUpMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnKeyDownMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyDownMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyDownMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnKeyUpMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyUpMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyUpMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnKeyPressMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnKeyPressMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnKeyPressMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnErrorMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnErrorMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnErrorMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnFilterMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnFilterMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnFilterMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnApplyFilterMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnApplyFilterMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnApplyFilterMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnTimerMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnTimerMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnTimerMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnUndoMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUndoMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUndoMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnRecordExitMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnRecordExitMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnRecordExitMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeginBatchEditMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeginBatchEditMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeginBatchEditMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string UndoBatchEditMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UndoBatchEditMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UndoBatchEditMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeBeginTransactionMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeBeginTransactionMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeBeginTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterBeginTransactionMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterBeginTransactionMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterBeginTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeCommitTransactionMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeCommitTransactionMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeCommitTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterCommitTransactionMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterCommitTransactionMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterCommitTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string RollbackTransactionMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RollbackTransactionMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RollbackTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnConnectMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnConnectMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnConnectMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDisconnectMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDisconnectMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDisconnectMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string PivotTableChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PivotTableChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PivotTableChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string QueryMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "QueryMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "QueryMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeQueryMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeQueryMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeQueryMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string SelectionChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SelectionChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelectionChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CommandBeforeExecuteMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandBeforeExecuteMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandBeforeExecuteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CommandCheckedMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandCheckedMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandCheckedMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CommandEnabledMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandEnabledMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandEnabledMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CommandExecuteMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandExecuteMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandExecuteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DataSetChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSetChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSetChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeScreenTipMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeScreenTipMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeScreenTipMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterFinalRenderMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterFinalRenderMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterFinalRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterRenderMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterRenderMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterLayoutMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterLayoutMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterLayoutMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeRenderMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeRenderMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string MouseWheelMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MouseWheelMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MouseWheelMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ViewChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ViewChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DataChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool AllowLayoutView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowLayoutView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowLayoutView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public Int32 DatasheetAlternateBackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DatasheetAlternateBackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetAlternateBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public byte DisplayOnSharePointSite
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DisplayOnSharePointSite");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayOnSharePointSite", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public Int32 SplitFormSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SplitFormSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SplitFormSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi._Section get_Section(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi._Section>(this, "Section", NetOffice.AccessApi._Section.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public NetOffice.AccessApi._Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public string RibbonName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RibbonName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RibbonName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool FitToScreen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FitToScreen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FitToScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Section get_SectionOld(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "SectionOld", NetOffice.AccessApi.Section.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Alias for get_SectionOld
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 14,15,16), Redirect("get_SectionOld")]
		public NetOffice.AccessApi.Section SectionOld(object index)
		{
			return get_SectionOld(index);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Recalc()
		{
			 Factory.ExecuteMethod(this, "Recalc");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Requery()
		{
			 Factory.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Repaint()
		{
			 Factory.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		/// <param name="down">optional Int32 Down = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber, object right, object down)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber, right, down);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pageNumber">Int32 pageNumber</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pageNumber">Int32 pageNumber</param>
		/// <param name="right">optional Int32 Right = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void GoToPage(Int32 pageNumber, object right)
		{
			 Factory.ExecuteMethod(this, "GoToPage", pageNumber, right);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object _Evaluate(string bstrExpr, object[] ppsa)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object _Evaluate(string bstrExpr)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top, object width, object height)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left)
		{
			 Factory.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top, object width)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width);
		}

        #endregion

        #region IEnumerable<object> Member

        /// <summary>
        /// SupportByVersion Access, 12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Access, 12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}



