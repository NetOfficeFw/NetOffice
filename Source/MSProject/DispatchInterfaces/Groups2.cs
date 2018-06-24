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
	/// DispatchInterface Groups2 
	/// SupportByVersion MSProject, 11,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920620(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("11589059-69F0-11D2-B889-00C04FB90729")]
	public interface Groups2 : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.Group2>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.Group2 this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Project Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="fieldName">string fieldName</param>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Group2 Add(string name, string fieldName);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="newName">string newName</param>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Group2 Copy(string name, string newName);

        #endregion


        #region IEnumerable<NetOffice.MSProjectApi.Group2>

        /// <summary>
        /// SupportByVersion MSProject, 11,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 14)]
        new IEnumerator<NetOffice.MSProjectApi.Group2> GetEnumerator();

        #endregion
    }
}
