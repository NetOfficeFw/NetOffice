using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// Interface ADORecordsetConstruction 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000283-0000-0010-8000-00AA006D2EA4")]
	public interface ADORecordsetConstruction : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object Rowset { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 Chapter { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object RowPosition { get; set; }

		#endregion

	}
}
