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
	/// DispatchInterface Exceptions 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920590(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("9DBAEC97-ADA1-4488-8845-818E734F182E")]
	public interface Exceptions : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.Exception>
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
		NetOffice.MSProjectApi.Calendar Parent { get; }

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
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.Exception this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		/// <param name="monthDay">optional object monthDay</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month, object monthDay);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType type</param>
		/// <param name="start">object start</param>
		/// <param name="finish">optional object finish</param>
		/// <param name="occurrences">optional object occurrences</param>
		/// <param name="name">optional object name</param>
		/// <param name="period">optional object period</param>
		/// <param name="daysOfWeek">optional object daysOfWeek</param>
		/// <param name="monthPosition">optional object monthPosition</param>
		/// <param name="monthItem">optional object monthItem</param>
		/// <param name="month">optional object month</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Exception>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.Exception> GetEnumerator();

        #endregion
    }
}
