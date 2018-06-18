using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementNamespaceFactoryCallback 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F7FD-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementNamespaceFactoryCallback : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrNamespace">string bstrNamespace</param>
		/// <param name="bstrTagName">string bstrTagName</param>
		/// <param name="bstrAttrs">string bstrAttrs</param>
		/// <param name="pNamespace">NetOffice.MSHTMLApi.IElementNamespace pNamespace</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Resolve(string bstrNamespace, string bstrTagName, string bstrAttrs, NetOffice.MSHTMLApi.IElementNamespace pNamespace);

		#endregion
	}
}
