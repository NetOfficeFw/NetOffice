using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IDocumentSelector 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("30510462-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IDocumentSelector : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement querySelector(string v);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="v">string v</param>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IHTMLDOMChildrenCollection querySelectorAll(string v);

		#endregion
	}
}
