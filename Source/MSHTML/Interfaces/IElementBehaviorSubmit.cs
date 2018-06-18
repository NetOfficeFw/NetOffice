using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorSubmit 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F646-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorSubmit : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSubmitData">NetOffice.MSHTMLApi.IHTMLSubmitData pSubmitData</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetSubmitInfo(NetOffice.MSHTMLApi.IHTMLSubmitData pSubmitData);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 reset();

		#endregion
	}
}
