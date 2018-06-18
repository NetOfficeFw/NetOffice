using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLTextRangeMetrics2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3050F4A6-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLTextRangeMetrics2 : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLRectCollection getClientRects();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLRect getBoundingClientRect();

		#endregion
	}
}
