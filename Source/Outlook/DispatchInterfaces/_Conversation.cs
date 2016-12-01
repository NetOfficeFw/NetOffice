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
	/// DispatchInterface _Conversation 
	/// SupportByVersion Outlook, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Conversation : COMObject
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
                    _type = typeof(_Conversation);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Conversation(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869259.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
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
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868054.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
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
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866390.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
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
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869565.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
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
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869792.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public string ConversationID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConversationID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866231.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Table GetTable()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetTable", paramsArray);
			NetOffice.OutlookApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Table.LateBindingApiWrapperType) as NetOffice.OutlookApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868807.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.SimpleItems GetChildren(object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			object returnItem = Invoker.MethodReturn(this, "GetChildren", paramsArray);
			NetOffice.OutlookApi.SimpleItems newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.SimpleItems.LateBindingApiWrapperType) as NetOffice.OutlookApi.SimpleItems;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869780.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public object GetParent(object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			object returnItem = Invoker.MethodReturn(this, "GetParent", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866457.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.SimpleItems GetRootItems()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetRootItems", paramsArray);
			NetOffice.OutlookApi.SimpleItems newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.SimpleItems.LateBindingApiWrapperType) as NetOffice.OutlookApi.SimpleItems;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869225.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public string GetAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			object returnItem = Invoker.MethodReturn(this, "GetAlwaysAssignCategories", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867861.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation GetAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			object returnItem = Invoker.MethodReturn(this, "GetAlwaysDelete", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869753.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.MAPIFolder GetAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			object returnItem = Invoker.MethodReturn(this, "GetAlwaysMoveToFolder", paramsArray);
			NetOffice.OutlookApi.MAPIFolder newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi.MAPIFolder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867852.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void MarkAsRead()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MarkAsRead", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868412.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void MarkAsUnread()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MarkAsUnread", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868084.aspx
		/// </summary>
		/// <param name="categories">string Categories</param>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void SetAlwaysAssignCategories(string categories, NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(categories, store);
			Invoker.Method(this, "SetAlwaysAssignCategories", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869857.aspx
		/// </summary>
		/// <param name="alwaysDelete">NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation AlwaysDelete</param>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void SetAlwaysDelete(NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete, NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alwaysDelete, store);
			Invoker.Method(this, "SetAlwaysDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865038.aspx
		/// </summary>
		/// <param name="moveToFolder">NetOffice.OutlookApi.MAPIFolder MoveToFolder</param>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void SetAlwaysMoveToFolder(NetOffice.OutlookApi.MAPIFolder moveToFolder, NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(moveToFolder, store);
			Invoker.Method(this, "SetAlwaysMoveToFolder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860425.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void ClearAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			Invoker.Method(this, "ClearAlwaysAssignCategories", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869032.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void StopAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			Invoker.Method(this, "StopAlwaysDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863707.aspx
		/// </summary>
		/// <param name="store">NetOffice.OutlookApi._Store Store</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void StopAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(store);
			Invoker.Method(this, "StopAlwaysMoveToFolder", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}