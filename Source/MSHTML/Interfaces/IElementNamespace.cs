using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementNamespace 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F671-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementNamespace : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrTagName">string bstrTagName</param>
		/// <param name="lFlags">Int32 lFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AddTag(string bstrTagName, Int32 lFlags);

		#endregion
	}
}
