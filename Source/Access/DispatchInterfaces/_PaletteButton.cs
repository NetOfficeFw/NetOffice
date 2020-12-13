using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _PaletteButton 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _PaletteButton : COMObject
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
                    _type = typeof(_PaletteButton);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _PaletteButton(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _PaletteButton(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _PaletteButton(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		public object OldValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OldValue");
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
		public NetOffice.AccessApi.Children Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Children>(this, "Controls", NetOffice.AccessApi.Children.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string EventProcPrefix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EventProcPrefix");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EventProcPrefix", value);
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
		public byte ControlType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ControlType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 OptionValue
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OptionValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OptionValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ControlSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string DefaultValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ValidationRule
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ValidationRule");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ValidationRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ValidationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ValidationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ValidationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string StatusBarText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StatusBarText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StatusBarText", value);
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
		public byte DisplayWhen
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DisplayWhen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayWhen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Locked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Locked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Locked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool TripleState
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TripleState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TripleState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool TabStop
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TabStop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabStop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 TabIndex
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TabIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool HideDuplicates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideDuplicates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideDuplicates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Left
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Top
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
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
		public Int16 Height
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte BackStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BackStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
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
		public string ControlTipText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlTipText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlTipText", value);
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
		public Int16 ColumnWidth
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ColumnWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 ColumnOrder
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ColumnOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool ColumnHidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ColumnHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AutoLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool AddColon
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AddColon");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddColon", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 LabelX
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LabelX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 LabelY
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LabelY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte LabelAlign
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "LabelAlign");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Section
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Section");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Section", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ControlName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlName", value);
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
		public bool IsVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool InSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InSelection", value);
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
		public string OnEnter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnEnter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnEnter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnExit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnExit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnExit", value);
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
		public string OnEnterMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnEnterMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnEnterMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnExitMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnExitMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnExitMacro", value);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SizeToFit()
		{
			 Factory.ExecuteMethod(this, "SizeToFit");
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Goto()
		{
			 Factory.ExecuteMethod(this, "Goto");
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

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
