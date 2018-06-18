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
	[TypeId("9DD14141-F0A9-4692-8288-A6835F93DC8A")]
	public interface StartDriver : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ActualStartDrivers ActualStartDrivers { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.PredecessorDrivers PredecessorDrivers { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ChildDrivers ChildTaskDrivers { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.CalendarDrivers CalendarDrivers { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Task Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 Suggestions { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 Warnings { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSProjectApi.OverAllocatedAssignments get_OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_OverAllocatedAssignments
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_OverAllocatedAssignments")]
		NetOffice.MSProjectApi.OverAllocatedAssignments OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="finishDate">object finishDate</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_EffectiveDateDifference(object startDate, object finishDate);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateDifference
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="finishDate">object finishDate</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateDifference")]
		object EffectiveDateDifference(object startDate, object finishDate);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_EffectiveDateAdd(object date, object duration);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateAdd
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateAdd")]
		object EffectiveDateAdd(object date, object duration);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_EffectiveDateSubtract(object date, object duration);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateSubtract
		/// </summary>
		/// <param name="date">object date</param>
		/// <param name="duration">object duration</param>
		[SupportByVersion("MSProject", 11,14), Redirect("get_EffectiveDateSubtract")]
		object EffectiveDateSubtract(object date, object duration);

		#endregion

	}
}
