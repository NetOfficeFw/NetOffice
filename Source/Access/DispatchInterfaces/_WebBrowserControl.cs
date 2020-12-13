using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _WebBrowserControl 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _WebBrowserControl : NetOffice.OfficeApi.IAccessible
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
                    _type = typeof(_WebBrowserControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _WebBrowserControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _WebBrowserControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _WebBrowserControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OldValue"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Properties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Controls"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Hyperlink"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Value"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ControlSource"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.DisplayWhen"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Enabled"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ReadyState"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcWebBrowserState ReadyState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcWebBrowserState>(this, "ReadyState");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Progress"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 Progress
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Progress");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ScrollBars"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.Enums.AcWebBrowserScrollBars ScrollBars
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcWebBrowserScrollBars>(this, "ScrollBars");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ScrollTop"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 ScrollTop
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ScrollTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ScrollLeft"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int32 ScrollLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ScrollLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Transform"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string Transform
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Transform");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Transform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.LocationURL"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string LocationURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LocationURL");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Object"/> </remarks>
		[SupportByVersion("Access", 14,15,16), ProxyResult]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Left"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Top"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Height"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.SpecialEffect"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public Int16 SpecialEffect
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "SpecialEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderStyle"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderColor"/> </remarks>
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

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderWidth"/> </remarks>
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
		[SupportByVersion("Access", 14,15,16)]
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
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Section"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnMouseDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnMouseMove"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnMouseUp"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnKeyDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnKeyUp"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnKeyPress"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Name"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Layout"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.LeftPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.TopPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.RightPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BottomPadding"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineStyleLeft"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineStyleTop"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineStyleRight"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineStyleBottom"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineWidthLeft"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineWidthTop"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineWidthRight"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineWidthBottom"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineColor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.HorizontalAnchor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.VerticalAnchor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.LayoutID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnBeforeNavigate"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string OnBeforeNavigate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnBeforeNavigate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnBeforeNavigate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnDocumentComplete"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string OnDocumentComplete
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDocumentComplete");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDocumentComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnNavigateError"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string OnNavigateError
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnNavigateError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnNavigateError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnProgressChange"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string OnProgressChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnProgressChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnProgressChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.OnUpdated"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string OnUpdated
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUpdated");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUpdated", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnBeforeNavigateMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnBeforeNavigateMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnBeforeNavigateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDocumentCompleteMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDocumentCompleteMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDocumentCompleteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnNavigateErrorMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnNavigateErrorMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnNavigateErrorMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnProgressChangeMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnProgressChangeMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnProgressChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnUpdatedMacro
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUpdatedMacro");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUpdatedMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.TabStop"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.TabIndex"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Visible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.InSelection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ControlType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.BorderShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineThemeColorIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineTint"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.GridlineShade"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.StatusBarText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.ControlTipText"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.EventProcPrefix"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.HelpContextId"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Tag"/> </remarks>
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Undo"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.SizeToFit"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public void SizeToFit()
		{
			 Factory.ExecuteMethod(this, "SizeToFit");
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Requery"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.SetFocus"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.WebBrowserControl.Move"/> </remarks>
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
