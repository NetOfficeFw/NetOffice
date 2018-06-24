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
	/// DispatchInterface CodeMask 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920570(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("4CC10F2B-7DF1-413C-A44D-9AB35ADFD9AE")]
	public interface CodeMask : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.CodeMaskLevel>
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
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.CodeMaskLevel this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.OutlineCode Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="sequence">optional NetOffice.MSProjectApi.Enums.PjCustomOutlineCodeSequence Sequence = 0</param>
		/// <param name="length">optional object length</param>
		/// <param name="separator">optional object separator</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.CodeMaskLevel Add(object sequence, object length, object separator);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.CodeMaskLevel Add();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="sequence">optional NetOffice.MSProjectApi.Enums.PjCustomOutlineCodeSequence Sequence = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.CodeMaskLevel Add(object sequence);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="sequence">optional NetOffice.MSProjectApi.Enums.PjCustomOutlineCodeSequence Sequence = 0</param>
		/// <param name="length">optional object length</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.CodeMaskLevel Add(object sequence, object length);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.CodeMaskLevel>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.CodeMaskLevel> GetEnumerator();

        #endregion
    }
}
