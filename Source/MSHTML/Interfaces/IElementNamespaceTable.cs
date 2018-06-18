using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementNamespaceTable 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F670-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementNamespaceTable : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrNamespace">string bstrNamespace</param>
		/// <param name="bstrUrn">string bstrUrn</param>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pvarFactory">object pvarFactory</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddNamespace(string bstrNamespace, string bstrUrn, Int32 lFlags, object pvarFactory);

		#endregion
	}
}
