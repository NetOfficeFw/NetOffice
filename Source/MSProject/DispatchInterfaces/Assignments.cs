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
	/// DispatchInterface Assignments 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920549(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0C46-0000-0000-C000-000000000046")]
	public interface Assignments : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.Assignment>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Project Parent { get; }

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
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSProjectApi.Assignment get_UniqueID(Int32 index);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Alias for get_UniqueID
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,12,14), Redirect("get_UniqueID")]
		NetOffice.MSProjectApi.Assignment UniqueID(Int32 index);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.Assignment this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskID">optional object taskID</param>
		/// <param name="resourceID">optional object resourceID</param>
		/// <param name="units">optional object units</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Assignment Add(object taskID, object resourceID, object units);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Assignment Add();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskID">optional object taskID</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Assignment Add(object taskID);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskID">optional object taskID</param>
		/// <param name="resourceID">optional object resourceID</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Assignment Add(object taskID, object resourceID);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Assignment>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.Assignment> GetEnumerator();

        #endregion
    }
}
