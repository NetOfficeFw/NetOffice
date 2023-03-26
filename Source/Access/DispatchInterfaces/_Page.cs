using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Page 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Page : NetOffice.OfficeApi.IAccessible
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
                    _type = typeof(_Page);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Page(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Page(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Page(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Properties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Controls"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.EventProcPrefix"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.ControlType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Caption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Picture"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.PictureType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.PageIndex"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 PageIndex
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PageIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.StatusBarText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Visible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Enabled"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Left"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Top"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Height"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.ShortcutMenuBar"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.ControlTipText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.HelpContextId"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Section"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Tag"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.IsVisible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.InSelection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.OnClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.OnDblClick"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.OnMouseDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.OnMouseMove"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.OnMouseUp"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.PictureData"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Name"/> </remarks>
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.SizeToFit"/> </remarks>
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Requery"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.SetFocus"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void SetFocus()
		{
			 Factory.ExecuteMethod(this, "SetFocus");
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.Move"/> </remarks>
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

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Page.SetTabOrder"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public void SetTabOrder()
		{
			 Factory.ExecuteMethod(this, "SetTabOrder");
		}

		#endregion

		#pragma warning restore
	}
}
