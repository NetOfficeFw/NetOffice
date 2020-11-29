using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _CalendarSharing 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CalendarSharing : COMObject
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
                    _type = typeof(_CalendarSharing);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CalendarSharing(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CalendarSharing(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CalendarSharing(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.Application"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.Class"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.Session"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.Parent"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.CalendarDetail"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlCalendarDetail CalendarDetail
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlCalendarDetail>(this, "CalendarDetail");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CalendarDetail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.EndDate"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime EndDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "EndDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EndDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.Folder"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder Folder
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi.MAPIFolder>(this, "Folder");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.IncludeAttachments"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IncludeAttachments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeAttachments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeAttachments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.IncludePrivateDetails"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IncludePrivateDetails
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludePrivateDetails");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludePrivateDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.RestrictToWorkingHours"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool RestrictToWorkingHours
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RestrictToWorkingHours");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RestrictToWorkingHours", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.StartDate"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime StartDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "StartDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.IncludeWholeCalendar"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IncludeWholeCalendar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeWholeCalendar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeWholeCalendar", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.SaveAsICal"/> </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SaveAsICal(string path)
		{
			 Factory.ExecuteMethod(this, "SaveAsICal", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.CalendarSharing.ForwardAsICal"/> </remarks>
		/// <param name="mailFormat">NetOffice.OutlookApi.Enums.OlCalendarMailFormat mailFormat</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsICal(NetOffice.OutlookApi.Enums.OlCalendarMailFormat mailFormat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsICal", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType, mailFormat);
		}

		#endregion

		#pragma warning restore
	}
}
