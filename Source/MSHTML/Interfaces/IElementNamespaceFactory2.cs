using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementNamespaceFactory2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F805-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementNamespaceFactory2 : IElementNamespaceFactory
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pNamespace">NetOffice.MSHTMLApi.IElementNamespace pNamespace</param>
		/// <param name="bstrImplementation">string bstrImplementation</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateWithImplementation(NetOffice.MSHTMLApi.IElementNamespace pNamespace, string bstrImplementation);

		#endregion
	}
}
