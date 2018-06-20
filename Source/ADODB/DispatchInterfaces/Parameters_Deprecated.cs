using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Parameters_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "ADODB", 2.5), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("0000050D-0000-0010-8000-00AA006D2EA4")]
	public interface Parameters_Deprecated : _DynaCollection, IEnumerableProvider<NetOffice.ADODBApi._Parameter_Deprecated>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new Int32 Count { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ADODBApi._Parameter_Deprecated this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new void Refresh();

        #endregion

        #region IEnumerable<NetOffice.ADODBApi._Parameter_Deprecated>

        /// <summary>
        /// SupportByVersion ADODB, 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.5)]
        new IEnumerator<NetOffice.ADODBApi._Parameter_Deprecated> GetEnumerator();

        #endregion
    }
}
