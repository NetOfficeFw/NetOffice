using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// CoClass HTMLStorage 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("30510475-98B5-11CF-BB82-00AA00BDCE0B")]
 	public interface HTMLStorage : DispHTMLStorage
	{

	}
}
