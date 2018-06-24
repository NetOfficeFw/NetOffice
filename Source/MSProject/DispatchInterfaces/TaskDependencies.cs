using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface TaskDependencies 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920720(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("A7C6ACB1-58C6-11D2-AE5A-00C04FB6F701")]
	public interface TaskDependencies : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.TaskDependency>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.TaskDependency this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 Count { get; }

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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Type = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from, object type, object lag);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="from">NetOffice.MSProjectApi.Task from</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Type = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TaskDependency Add(NetOffice.MSProjectApi.Task from, object type);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.TaskDependency>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.TaskDependency> GetEnumerator();

        #endregion
    }
}
