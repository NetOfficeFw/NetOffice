using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// CoClass HTMLNavigator 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("FECEAAA6-8405-11CF-8BA1-00AA00476DA6")]
 	public interface HTMLNavigator : DispHTMLNavigator
	{

	}
}
