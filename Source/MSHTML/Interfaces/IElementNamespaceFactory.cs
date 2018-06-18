using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementNamespaceFactory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F672-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementNamespaceFactory : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pNamespace">NetOffice.MSHTMLApi.IElementNamespace pNamespace</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 create(NetOffice.MSHTMLApi.IElementNamespace pNamespace);

		#endregion
	}
}
