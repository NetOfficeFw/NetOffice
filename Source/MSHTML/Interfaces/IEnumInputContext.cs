using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumInputContext 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("09B5EAB0-F997-11D1-93D4-0060B067B86E")]
	public interface IEnumInputContext : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumInputContext ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Clone(out NetOffice.MSHTMLApi.IEnumInputContext ppEnum);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ulCount">Int32 ulCount</param>
		/// <param name="rgInputContext">Int32 rgInputContext</param>
		/// <param name="pcFetched">Int32 pcFetched</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Next(Int32 ulCount, out Int32 rgInputContext, out Int32 pcFetched);

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
