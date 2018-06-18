using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumRegisterWordA 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("08C03412-F96B-11D0-A475-00AA006BCC59")]
	public interface IEnumRegisterWordA : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumRegisterWordA ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Clone(out NetOffice.MSHTMLApi.IEnumRegisterWordA ppEnum);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ulCount">Int32 ulCount</param>
		/// <param name="rgRegisterWord">__MIDL___MIDL_itf_mshtml_0001_0042_0001 rgRegisterWord</param>
		/// <param name="pcFetched">Int32 pcFetched</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Next(Int32 ulCount, out __MIDL___MIDL_itf_mshtml_0001_0042_0001 rgRegisterWord, out Int32 pcFetched);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 reset();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ulCount">Int32 ulCount</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Skip(Int32 ulCount);

		#endregion
	}
}
