using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Image 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Image : COMObject
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
                    _type = typeof(_Image);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Image(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Image(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Image(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OldValue"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Properties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Controls"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Hyperlink"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.EventProcPrefix"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ControlType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Picture"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.PictureType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.PictureData"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.SizeMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte SizeMode
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "SizeMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.PictureAlignment"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.PictureTiling"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.HyperlinkAddress"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.HyperlinkSubAddress"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Visible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.DisplayWhen"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Left"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Top"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Height"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BackStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BackColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.SpecialEffect"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OldBorderStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderWidth"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ControlTipText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.HelpContextId"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Section"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ObjectPalette"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.IsVisible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OnClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OnDblClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OnMouseDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OnMouseMove"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.OnMouseUp"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ShortcutMenuBar"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ImageHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 ImageHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ImageHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ImageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.ImageWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 ImageWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ImageWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ImageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.InSelection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Tag"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Name"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Layout"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.LeftPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.TopPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.RightPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BottomPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineStyleLeft"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineStyleTop"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineStyleRight"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineStyleBottom"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineWidthLeft"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineWidthTop"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineWidthRight"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineWidthBottom"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.HorizontalAnchor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.VerticalAnchor"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.LayoutID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BackThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BackTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BackShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.BorderShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.GridlineShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.SizeToFit"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SizeToFit()
		{
			 Factory.ExecuteMethod(this, "SizeToFit");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Requery"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.SetFocus"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Image.Move"/> </remarks>
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
