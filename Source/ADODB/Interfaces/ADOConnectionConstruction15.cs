using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// Interface ADOConnectionConstruction15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("00000516-0000-0010-8000-00AA006D2EA4")]
	public interface ADOConnectionConstruction15 : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object DSO { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		object Session { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="pDSO">object pDSO</param>
		/// <param name="pSession">object pSession</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 WrapDSOandSession(object pDSO, object pSession);

		#endregion
	}
}
