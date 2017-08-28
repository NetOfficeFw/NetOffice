using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface RecurrencePattern 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863458.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class RecurrencePattern : COMObject
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
                    _type = typeof(RecurrencePattern);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public RecurrencePattern(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public RecurrencePattern(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public RecurrencePattern(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869472.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865875.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869907.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867159.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869388.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 DayOfMonth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DayOfMonth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DayOfMonth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866785.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDaysOfWeek DayOfWeekMask
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDaysOfWeek>(this, "DayOfWeekMask");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DayOfWeekMask", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867680.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Duration
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Duration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866939.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime EndTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "EndTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EndTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869544.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Exceptions Exceptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Exceptions>(this, "Exceptions", NetOffice.OutlookApi.Exceptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863347.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Instance
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Instance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Instance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869589.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Interval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Interval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Interval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861331.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 MonthOfYear
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MonthOfYear");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MonthOfYear", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864419.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool NoEndDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NoEndDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoEndDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868430.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Occurrences
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Occurrences");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Occurrences", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861049.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime PatternEndDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "PatternEndDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PatternEndDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862199.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime PatternStartDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "PatternStartDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PatternStartDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868812.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlRecurrenceType RecurrenceType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRecurrenceType>(this, "RecurrenceType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RecurrenceType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868917.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool Regenerate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Regenerate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Regenerate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865086.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime StartTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "StartTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartTime", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862727.aspx </remarks>
		/// <param name="startDate">DateTime startDate</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.AppointmentItem GetOccurrence(DateTime startDate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AppointmentItem>(this, "GetOccurrence", NetOffice.OutlookApi.AppointmentItem.LateBindingApiWrapperType, startDate);
		}

		#endregion

		#pragma warning restore
	}
}
