using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _Conversation 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Conversation : COMObject, NetOffice.OutlookApi._Conversation
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OutlookApi._Conversation);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Conversation() : base()
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
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868054.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866390.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869565.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869792.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ConversationID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationID");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866231.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Table GetTable()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", typeof(NetOffice.OutlookApi.Table));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868807.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.SimpleItems GetChildren(object item)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SimpleItems>(this, "GetChildren", typeof(NetOffice.OutlookApi.SimpleItems), item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869780.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual object GetParent(object item)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetParent", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866457.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.SimpleItems GetRootItems()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SimpleItems>(this, "GetRootItems", typeof(NetOffice.OutlookApi.SimpleItems));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869225.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string GetAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAlwaysAssignCategories", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867861.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation GetAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation>(this, "GetAlwaysDelete", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869753.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi.MAPIFolder GetAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "GetAlwaysMoveToFolder", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867852.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void MarkAsRead()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAsRead");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868412.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void MarkAsUnread()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAsUnread");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868084.aspx </remarks>
		/// <param name="categories">string categories</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SetAlwaysAssignCategories(string categories, NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAlwaysAssignCategories", categories, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869857.aspx </remarks>
		/// <param name="alwaysDelete">NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SetAlwaysDelete(NetOffice.OutlookApi.Enums.OlAlwaysDeleteConversation alwaysDelete, NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAlwaysDelete", alwaysDelete, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865038.aspx </remarks>
		/// <param name="moveToFolder">NetOffice.OutlookApi.MAPIFolder moveToFolder</param>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SetAlwaysMoveToFolder(NetOffice.OutlookApi.MAPIFolder moveToFolder, NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAlwaysMoveToFolder", moveToFolder, store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860425.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void ClearAlwaysAssignCategories(NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearAlwaysAssignCategories", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869032.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void StopAlwaysDelete(NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StopAlwaysDelete", store);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863707.aspx </remarks>
		/// <param name="store">NetOffice.OutlookApi._Store store</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void StopAlwaysMoveToFolder(NetOffice.OutlookApi._Store store)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StopAlwaysMoveToFolder", store);
		}

		#endregion

		#pragma warning restore
	}
}


