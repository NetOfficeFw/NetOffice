using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsEnum)]
	public enum _styleWidowsOrphans
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>-2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleWidowsOrphansNotSet = -2147483647,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleWidowsOrphans_Max = 2147483647
	}
}