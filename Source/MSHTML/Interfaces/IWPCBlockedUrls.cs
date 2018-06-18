using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IWPCBlockedUrls 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("30510413-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IWPCBlockedUrls : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pdwCount">Int32 pdwCount</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCount(out Int32 pdwCount);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwIdx">Int32 dwIdx</param>
		/// <param name="pbstrUrl">string pbstrUrl</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetUrl(Int32 dwIdx, out string pbstrUrl);

		#endregion
	}
}
