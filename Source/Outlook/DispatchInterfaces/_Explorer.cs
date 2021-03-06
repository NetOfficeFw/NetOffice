﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Explorer 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Explorer : COMObject
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
                    _type = typeof(_Explorer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Explorer(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Explorer(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Application"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Class"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Session"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Parent"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.CurrentFolder"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder CurrentFolder
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi.MAPIFolder>(this, "CurrentFolder");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "CurrentFolder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Caption"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.CurrentView"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object CurrentView
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CurrentView");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CurrentView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Height"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Left"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Panes"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Panes Panes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Panes>(this, "Panes", NetOffice.OutlookApi.Panes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Selection"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Selection Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Selection>(this, "Selection", NetOffice.OutlookApi.Selection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Top"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Width"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.WindowState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlWindowState WindowState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlWindowState>(this, "WindowState");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Views
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Views");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.HTMLDocument"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		public object HTMLDocument
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HTMLDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.NavigationPane"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NavigationPane NavigationPane
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NavigationPane>(this, "NavigationPane");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.AccountSelector"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._AccountSelector AccountSelector
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._AccountSelector>(this, "AccountSelector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.AttachmentSelection"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._AttachmentSelection AttachmentSelection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._AttachmentSelection>(this, "AttachmentSelection");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.explorer.activeinlineresponse"/> </remarks>
		[SupportByVersion("Outlook", 15, 16), ProxyResult]
		public object ActiveInlineResponse
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveInlineResponse");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.explorer.activeinlineresponsewordeditor"/> </remarks>
		[SupportByVersion("Outlook", 15, 16), ProxyResult]
		public object ActiveInlineResponseWordEditor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveInlineResponseWordEditor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Close(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Display"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Activate(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.IsPaneVisible"/> </remarks>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane pane</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool IsPaneVisible(NetOffice.OutlookApi.Enums.OlPane pane)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsPaneVisible", pane);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.ShowPane"/> </remarks>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane pane</param>
		/// <param name="visible">bool visible</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void ShowPane(NetOffice.OutlookApi.Enums.OlPane pane, bool visible)
		{
			 Factory.ExecuteMethod(this, "ShowPane", pane, visible);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void SelectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			 Factory.ExecuteMethod(this, "SelectFolder", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void DeselectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			 Factory.ExecuteMethod(this, "DeselectFolder", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public bool IsFolderSelected(NetOffice.OutlookApi.MAPIFolder folder)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsFolderSelected", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.ClearSearch"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ClearSearch()
		{
			 Factory.ExecuteMethod(this, "ClearSearch");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.Search"/> </remarks>
		/// <param name="query">string query</param>
		/// <param name="searchScope">NetOffice.OutlookApi.Enums.OlSearchScope searchScope</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Search(string query, NetOffice.OutlookApi.Enums.OlSearchScope searchScope)
		{
			 Factory.ExecuteMethod(this, "Search", query, searchScope);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.IsItemSelectableInView"/> </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool IsItemSelectableInView(object item)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsItemSelectableInView", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.AddToSelection"/> </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void AddToSelection(object item)
		{
			 Factory.ExecuteMethod(this, "AddToSelection", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.RemoveFromSelection"/> </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void RemoveFromSelection(object item)
		{
			 Factory.ExecuteMethod(this, "RemoveFromSelection", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.SelectAllItems"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public void SelectAllItems()
		{
			 Factory.ExecuteMethod(this, "SelectAllItems");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Explorer.ClearSelection"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public void ClearSelection()
		{
			 Factory.ExecuteMethod(this, "ClearSelection");
		}

		#endregion

		#pragma warning restore
	}
}
