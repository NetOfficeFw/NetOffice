using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface StartDriver 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920699(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class StartDriver : COMObject
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
                    _type = typeof(StartDriver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public StartDriver(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public StartDriver(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ActualStartDrivers ActualStartDrivers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ActualStartDrivers>(this, "ActualStartDrivers", NetOffice.MSProjectApi.ActualStartDrivers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.PredecessorDrivers PredecessorDrivers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.PredecessorDrivers>(this, "PredecessorDrivers", NetOffice.MSProjectApi.PredecessorDrivers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ChildDrivers ChildTaskDrivers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ChildDrivers>(this, "ChildTaskDrivers", NetOffice.MSProjectApi.ChildDrivers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.CalendarDrivers CalendarDrivers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.CalendarDrivers>(this, "CalendarDrivers", NetOffice.MSProjectApi.CalendarDrivers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Task Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "Parent", NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 Suggestions
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Suggestions");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 Warnings
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Warnings");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSProjectApi.OverAllocatedAssignments get_OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.OverAllocatedAssignments>(this, "OverAllocatedAssignments", NetOffice.MSProjectApi.OverAllocatedAssignments.LateBindingApiWrapperType, overallocationType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_OverAllocatedAssignments
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_OverAllocatedAssignments")]
		public NetOffice.MSProjectApi.OverAllocatedAssignments OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
		{
			return get_OverAllocatedAssignments(overallocationType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="finishDate">object finishDate</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateDifference(object startDate, object finishDate)
		{
			return Factory.ExecuteVariantPropertyGet(this, "EffectiveDateDifference", startDate, finishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateDifference
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="finishDate">object finishDate</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateDifference")]
		public object EffectiveDateDifference(object startDate, object finishDate)
		{
			return get_EffectiveDateDifference(startDate, finishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateAdd(object date, object duration)
		{
			return Factory.ExecuteVariantPropertyGet(this, "EffectiveDateAdd", date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateAdd
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateAdd")]
		public object EffectiveDateAdd(object date, object duration)
		{
			return get_EffectiveDateAdd(date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateSubtract(object date, object duration)
		{
			return Factory.ExecuteVariantPropertyGet(this, "EffectiveDateSubtract", date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateSubtract
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateSubtract")]
		public object EffectiveDateSubtract(object date, object duration)
		{
			return get_EffectiveDateSubtract(date, duration);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
