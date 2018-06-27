using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface StartDriver 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920699(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class StartDriver : COMObject, NetOffice.MSProjectApi.StartDriver
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
                    _contractType = typeof(NetOffice.MSProjectApi.StartDriver);
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
                    _type = typeof(StartDriver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public StartDriver() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ActualStartDrivers ActualStartDrivers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ActualStartDrivers>(this, "ActualStartDrivers", typeof(NetOffice.MSProjectApi.ActualStartDrivers));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.PredecessorDrivers PredecessorDrivers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.PredecessorDrivers>(this, "PredecessorDrivers", typeof(NetOffice.MSProjectApi.PredecessorDrivers));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ChildDrivers ChildTaskDrivers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ChildDrivers>(this, "ChildTaskDrivers", typeof(NetOffice.MSProjectApi.ChildDrivers));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.CalendarDrivers CalendarDrivers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.CalendarDrivers>(this, "CalendarDrivers", typeof(NetOffice.MSProjectApi.CalendarDrivers));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Task Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "Parent", typeof(NetOffice.MSProjectApi.Task));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", typeof(NetOffice.MSProjectApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual Int32 Suggestions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Suggestions");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual Int32 Warnings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Warnings");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSProjectApi.OverAllocatedAssignments get_OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.OverAllocatedAssignments>(this, "OverAllocatedAssignments", typeof(NetOffice.MSProjectApi.OverAllocatedAssignments), overallocationType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_OverAllocatedAssignments
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_OverAllocatedAssignments")]
		public virtual NetOffice.MSProjectApi.OverAllocatedAssignments OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
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
		public virtual object get_EffectiveDateDifference(object startDate, object finishDate)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EffectiveDateDifference", startDate, finishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateDifference
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="finishDate">object finishDate</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateDifference")]
		public virtual object EffectiveDateDifference(object startDate, object finishDate)
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
		public virtual object get_EffectiveDateAdd(object date, object duration)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EffectiveDateAdd", date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateAdd
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateAdd")]
		public virtual object EffectiveDateAdd(object date, object duration)
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
		public virtual object get_EffectiveDateSubtract(object date, object duration)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EffectiveDateSubtract", date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateSubtract
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateSubtract")]
		public virtual object EffectiveDateSubtract(object date, object duration)
		{
			return get_EffectiveDateSubtract(date, duration);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


