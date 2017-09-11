using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Conversation 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Conversation : COMObject
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
                    _type = typeof(_Conversation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Conversation(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Conversation(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869259.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868054.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866390.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869565.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869792.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public string ConversationID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationID");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866231.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Table GetTable()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", NetOffice.OutlookApi.Table.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868807.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.SimpleItems GetChildren(object item)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SimpleItems>(this, "GetChildren", NetOffice.OutlookApi.SimpleItems.LateBindingApiWrapperType, item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869780.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public object GetParent(object item)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetParent", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866457.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.SimpleItems GetRootItems()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SimpleItems>(this, "GetRootItems", NetOffice.OutlookApi.SimpleItems.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869225.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public string GetAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAlwaysAssignCategories", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867861.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation GetAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation>(this, "GetAlwaysDelete", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869753.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder GetAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetAlwaysMoveToFolder", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867852.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public void MarkAsRead()
		{
			 Factory.ExecuteMethod(this, "MarkAsRead");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868412.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public void MarkAsUnread()
		{
			 Factory.ExecuteMethod(this, "MarkAsUnread");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868084.aspx </remarks>
		/// <param name="categories">string categories</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void SetAlwaysAssignCategories(string categories, NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "SetAlwaysAssignCategories", categories, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869857.aspx </remarks>
		/// <param name="alwaysDelete">NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void SetAlwaysDelete(NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete, NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "SetAlwaysDelete", alwaysDelete, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865038.aspx </remarks>
		/// <param name="moveToFolder">NetOffice.OutlookApi.MAPIFolder moveToFolder</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void SetAlwaysMoveToFolder(NetOffice.OutlookApi.MAPIFolder moveToFolder, NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "SetAlwaysMoveToFolder", moveToFolder, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860425.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void ClearAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "ClearAlwaysAssignCategories", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869032.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void StopAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "StopAlwaysDelete", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863707.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void StopAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			 Factory.ExecuteMethod(this, "StopAlwaysMoveToFolder", store);
		}

		#endregion

		#pragma warning restore
	}
}
