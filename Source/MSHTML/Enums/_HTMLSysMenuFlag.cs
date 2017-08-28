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
	public enum _HTMLSysMenuFlag
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLSysMenuFlagNo = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLSysMenuFlagYes = 524288,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLSysMenuFlag_Max = 2147483647
	}
}