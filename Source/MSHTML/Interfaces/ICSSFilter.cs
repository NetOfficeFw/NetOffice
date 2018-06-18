using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ICSSFilter 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F3EC-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface ICSSFilter : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSink">NetOffice.MSHTMLApi.ICSSFilterSite pSink</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetSite(NetOffice.MSHTMLApi.ICSSFilterSite pSink);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 OnAmbientPropertyChange(Int32 dispid);

		#endregion
	}
}
