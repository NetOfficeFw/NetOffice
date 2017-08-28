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
	public enum _styleZIndex
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>-2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleZIndexAuto = -2147483647,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleZIndex_Max = 2147483647
	}
}