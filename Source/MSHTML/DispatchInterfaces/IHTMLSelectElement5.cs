using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLSelectElement5 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("3051049D-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLSelectElement5 : IHTMLSelectElement4
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pElem">NetOffice.MSHTMLApi.IHTMLOptionElement pElem</param>
		/// <param name="pvarBefore">object pvarBefore</param>
		[SupportByVersion("MSHTML", 4)]
		void add(NetOffice.MSHTMLApi.IHTMLOptionElement pElem, object pvarBefore);

		#endregion
	}
}
