using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _Explorer 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Explorer : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Explorer(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869277.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OutlookApi._Application newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861302.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Class", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlObjectClass)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864406.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Session", paramsArray);
				NetOffice.OutlookApi._NameSpace newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863296.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866719.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder CurrentFolder
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentFolder", paramsArray);
				NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CurrentFolder", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866079.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Caption", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861617.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object CurrentView
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentView", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CurrentView", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868816.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Height", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867221.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Left", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868726.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Panes Panes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Panes", paramsArray);
				NetOffice.OutlookApi.Panes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Panes.LateBindingApiWrapperType) as NetOffice.OutlookApi.Panes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861267.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Selection Selection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selection", paramsArray);
				NetOffice.OutlookApi.Selection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.Selection.LateBindingApiWrapperType) as NetOffice.OutlookApi.Selection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869921.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Top", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867095.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Width", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866771.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlWindowState WindowState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WindowState", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlWindowState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WindowState", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Views
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Views", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869477.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public object HTMLDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLDocument", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868233.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi._NavigationPane NavigationPane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NavigationPane", paramsArray);
				NetOffice.OutlookApi._NavigationPane newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NavigationPane;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865605.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi._AccountSelector AccountSelector
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AccountSelector", paramsArray);
				NetOffice.OutlookApi._AccountSelector newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._AccountSelector;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869315.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi._AttachmentSelection AttachmentSelection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AttachmentSelection", paramsArray);
				NetOffice.OutlookApi._AttachmentSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._AttachmentSelection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231535.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 15, 16)]
		public object ActiveInlineResponse
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveInlineResponse", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230742.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 15, 16)]
		public object ActiveInlineResponseWordEditor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveInlineResponseWordEditor", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869525.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863904.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Display", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865063.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869316.aspx
		/// </summary>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane Pane</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public bool IsPaneVisible(NetOffice.OutlookApi.Enums.OlPane pane)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pane);
			object returnItem = Invoker.MethodReturn(this, "IsPaneVisible", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863893.aspx
		/// </summary>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane Pane</param>
		/// <param name="visible">bool Visible</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void ShowPane(NetOffice.OutlookApi.Enums.OlPane pane, bool visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pane, visible);
			Invoker.Method(this, "ShowPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder Folder</param>
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		public void SelectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(folder);
			Invoker.Method(this, "SelectFolder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder Folder</param>
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		public void DeselectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(folder);
			Invoker.Method(this, "DeselectFolder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder Folder</param>
		[SupportByVersionAttribute("Outlook", 11,12,14,15,16)]
		public bool IsFolderSelected(NetOffice.OutlookApi.MAPIFolder folder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(folder);
			object returnItem = Invoker.MethodReturn(this, "IsFolderSelected", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865861.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public void ClearSearch()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearSearch", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869309.aspx
		/// </summary>
		/// <param name="query">string Query</param>
		/// <param name="searchScope">NetOffice.OutlookApi.Enums.OlSearchScope SearchScope</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public void Search(string query, NetOffice.OutlookApi.Enums.OlSearchScope searchScope)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(query, searchScope);
			Invoker.Method(this, "Search", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868297.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public bool IsItemSelectableInView(object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			object returnItem = Invoker.MethodReturn(this, "IsItemSelectableInView", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868732.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void AddToSelection(object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			Invoker.Method(this, "AddToSelection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869912.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void RemoveFromSelection(object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			Invoker.Method(this, "RemoveFromSelection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860458.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void SelectAllItems()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAllItems", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862680.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void ClearSelection()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearSelection", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}