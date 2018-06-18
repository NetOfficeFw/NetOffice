using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLEditHost 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F6A0-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IHTMLEditHost : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="prcNew">tagRECT prcNew</param>
		/// <param name="eHandle">NetOffice.MSHTMLApi.Enums._ELEMENT_CORNER eHandle</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SnapRect(NetOffice.MSHTMLApi.IHTMLElement pIElement, tagRECT prcNew, NetOffice.MSHTMLApi.Enums._ELEMENT_CORNER eHandle);

		#endregion
	}
}
