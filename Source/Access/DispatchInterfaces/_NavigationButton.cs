using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _NavigationButton 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _NavigationButton : NetOffice.OfficeApi.IAccessible
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
                    _type = typeof(_NavigationButton);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _NavigationButton(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _NavigationButton(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _NavigationButton(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Application"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Parent"/> </remarks>
		[SupportByVersion("Access", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OldValue"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public object OldValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OldValue");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Properties"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", NetOffice.AccessApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Controls"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Children Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Children>(this, "Controls", NetOffice.AccessApi.Children.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Hyperlink"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._Hyperlink Hyperlink
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Hyperlink>(this, "Hyperlink");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.EventProcPrefix"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ControlType"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Caption"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Picture"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PictureType"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PictureData"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Transparent"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool Transparent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Transparent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Transparent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.AutoRepeat"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool AutoRepeat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoRepeat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoRepeat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.StatusBarText"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnPush"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnPush
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnPush");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnPush", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HyperlinkAddress"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string HyperlinkAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HyperlinkSubAddress"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string HyperlinkSubAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkSubAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkSubAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Visible"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.DisplayWhen"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Enabled"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.TabStop"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.TabIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Left"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 Left
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Top"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 Top
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Width"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Height"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ForeColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontName"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontSize"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 FontSize
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontWeight"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 FontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontItalic"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool FontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontUnderline"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool FontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte TextFontCharSet
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TextFontCharSet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFontCharSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.FontBold"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 FontBold
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ShortcutMenuBar"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ControlTipText"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HelpContextId"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.AutoLabel"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.AddColon"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.LabelX"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.LabelY"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.LabelAlign"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Section"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Tag"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ObjectPalette"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ObjectPalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ObjectPalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ObjectPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.IsVisible"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.InSelection"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnEnter"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnExit"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnGotFocus"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnLostFocus"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnClick"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnDblClick"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnMouseDown"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnMouseMove"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnMouseUp"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnKeyDown"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnKeyUp"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.OnKeyPress"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ReadingOrder"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte ReadingOrder
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "ReadingOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Name"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Layout"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcLayoutType Layout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcLayoutType>(this, "Layout");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.LeftPadding"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 LeftPadding
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LeftPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.TopPadding"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 TopPadding
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TopPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.RightPadding"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 RightPadding
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RightPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BottomPadding"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 BottomPadding
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "BottomPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineStyleLeft"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineStyleLeft
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineStyleLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineStyleLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineStyleTop"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineStyleTop
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineStyleTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineStyleTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineStyleRight"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineStyleRight
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineStyleRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineStyleRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineStyleBottom"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineStyleBottom
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineStyleBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineStyleBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineWidthLeft"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineWidthLeft
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineWidthLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineWidthLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineWidthTop"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineWidthTop
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineWidthTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineWidthTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineWidthRight"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineWidthRight
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineWidthRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineWidthRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineWidthBottom"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte GridlineWidthBottom
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GridlineWidthBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineWidthBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 GridlineColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridlineColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HorizontalAnchor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcHorizontalAnchor>(this, "HorizontalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.VerticalAnchor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcVerticalAnchor>(this, "VerticalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.LayoutID"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 LayoutID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LayoutID");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BackStyle"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.CursorOnHover"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcCursorOnHover CursorOnHover
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcCursorOnHover>(this, "CursorOnHover");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CursorOnHover", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PictureCaptionArrangement"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcPictureCaptionArrangement PictureCaptionArrangement
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPictureCaptionArrangement>(this, "PictureCaptionArrangement");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureCaptionArrangement", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Alignment"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte Alignment
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Alignment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Alignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Target
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Target");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Target", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ForeThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 ForeThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ForeTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single ForeTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ForeTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ForeShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single ForeShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ForeShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ThemeFontIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 ThemeFontIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ThemeFontIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ThemeFontIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 GridlineThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridlineThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single GridlineTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridlineTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.GridlineShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single GridlineShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridlineShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.NavigationTargetName"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string NavigationTargetName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NavigationTargetName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NavigationTargetName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.NavigationWhereClause"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string NavigationWhereClause
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NavigationWhereClause");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NavigationWhereClause", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.ParentTab"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.NavigationButton ParentTab
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.NavigationButton>(this, "ParentTab", NetOffice.AccessApi.NavigationButton.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BackThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 BackThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BackTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single BackTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BackTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BackShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single BackShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BackShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 BorderThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BorderThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single BorderTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BorderTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single BorderShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BorderShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderWidth"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public byte BorderWidth
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BorderWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderStyle"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 HoverColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HoverColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 HoverThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HoverThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single HoverTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "HoverTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single HoverShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "HoverShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverForeColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 HoverForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HoverForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverForeThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 HoverForeThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HoverForeThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverForeThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverForeTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single HoverForeTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "HoverForeTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverForeTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.HoverForeShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single HoverForeShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "HoverForeShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverForeShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 PressedColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PressedColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 PressedThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PressedThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single PressedTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PressedTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single PressedShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PressedShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedForeColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 PressedForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PressedForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedForeThemeColorIndex"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 PressedForeThemeColorIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PressedForeThemeColorIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedForeThemeColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedForeTint"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single PressedForeTint
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PressedForeTint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedForeTint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.PressedForeShade"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Single PressedForeShade
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PressedForeShade");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PressedForeShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Shape"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Shape
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Shape");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Shape", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Bevel"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Bevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Bevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Bevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Glow"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Glow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Glow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Glow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Shadow"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Shadow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Shadow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.SoftEdges"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 SoftEdges
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SoftEdges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SoftEdges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.QuickStyle"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 QuickStyle
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "QuickStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "QuickStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.navigationbutton.quickstylemask"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 QuickStyleMask
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "QuickStyleMask");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "QuickStyleMask", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Gradient"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Gradient
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Gradient");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Gradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BackColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.BorderColor"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 BorderColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BorderColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderColor", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.SizeToFit"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public void SizeToFit()
		{
			 Factory.ExecuteMethod(this, "SizeToFit");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Requery"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public void Requery()
		{
			 Factory.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public void Goto()
		{
			 Factory.ExecuteMethod(this, "Goto");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.SetFocus"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public object _Evaluate(string bstrExpr)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Move"/> </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 14,15,16)]
		public void Move(object left, object top, object width, object height)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Move"/> </remarks>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public void Move(object left)
		{
			 Factory.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Move"/> </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public void Move(object left, object top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.NavigationButton.Move"/> </remarks>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public void Move(object left, object top, object width)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
