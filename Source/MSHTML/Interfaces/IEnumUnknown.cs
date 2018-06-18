using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumUnknown 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000100-0000-0000-C000-000000000046")]
	public interface IEnumUnknown : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		/// <param name="rgelt">object rgelt</param>
		/// <param name="pceltFetched">Int32 pceltFetched</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RemoteNext(Int32 celt, out object rgelt, out Int32 pceltFetched);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Skip(Int32 celt);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 reset();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumUnknown ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Clone(out NetOffice.MSHTMLApi.IEnumUnknown ppEnum);

		#endregion
	}
}
