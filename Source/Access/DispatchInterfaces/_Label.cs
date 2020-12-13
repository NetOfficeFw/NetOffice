using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Label 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Label : NetOffice.OfficeApi.IAccessible
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
                    _type = typeof(_Label);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Label(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Label(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Label(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Properties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Hyperlink"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._Hyperlink Hyperlink
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Hyperlink>(this, "Hyperlink");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.EventProcPrefix"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ControlType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Caption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.HyperlinkAddress"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.HyperlinkSubAddress"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Visible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.DisplayWhen"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Vertical"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Vertical
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Vertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Vertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Left"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Top"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Height"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BackStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BackColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.SpecialEffect"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte SpecialEffect
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "SpecialEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OldBorderStyle"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte OldBorderStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "OldBorderStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OldBorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte BorderLineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "BorderLineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ForeColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontName"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontSize"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontWeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontItalic"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontUnderline"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.TextAlign"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte TextAlign
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "TextAlign");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.FontBold"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ControlTipText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.HelpContextId"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Section"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.IsVisible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.InSelection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OnClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OnDblClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OnMouseDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OnMouseMove"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.OnMouseUp"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ShortcutMenuBar"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Tag"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ReadingOrder"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.NumeralShapes"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte NumeralShapes
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "NumeralShapes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumeralShapes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Name"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.LeftMargin"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 LeftMargin
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LeftMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.TopMargin"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 TopMargin
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "TopMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.LineSpacing"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 LineSpacing
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LineSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.RightMargin"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 RightMargin
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "RightMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BottomMargin"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 BottomMargin
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "BottomMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.SmartTags"/> </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._SmartTags SmartTags
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._SmartTags>(this, "SmartTags");
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Layout"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcLayoutType Layout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcLayoutType>(this, "Layout");
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.LeftPadding"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.TopPadding"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.RightPadding"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BottomPadding"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineStyleLeft"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineStyleTop"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineStyleRight"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineStyleBottom"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineWidthLeft"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineWidthTop"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineWidthRight"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineWidthBottom"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineColor"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.HorizontalAnchor"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.VerticalAnchor"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.LayoutID"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BackThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BackTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BackShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.BorderShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ForeThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ForeTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ForeShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.ThemeFontIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.GridlineShade"/> </remarks>
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.SizeToFit"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SizeToFit()
		{
			 Factory.ExecuteMethod(this, "SizeToFit");
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
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr, ppsa);
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
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr);
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
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Label.Move"/> </remarks>
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
