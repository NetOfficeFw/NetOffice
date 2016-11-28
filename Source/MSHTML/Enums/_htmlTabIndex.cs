using System;
using NetOffice;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum _htmlTabIndex
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>-32768</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlTabIndexNotSet = -32768,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlTabIndex_Max = 2147483647
	}
}