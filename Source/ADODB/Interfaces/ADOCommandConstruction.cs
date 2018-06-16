using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// Interface ADOCommandConstruction 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000517-0000-0010-8000-00AA006D2EA4")]
	public interface ADOCommandConstruction : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object OLEDBCommand { get; set; }

		#endregion

	}
}
