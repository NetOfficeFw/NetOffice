using System;
using NetOffice;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum _HTMLBorder
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLBorderNone = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLBorderThick = 262144,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4194304</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLBorderDialog = 4194304,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>8388608</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLBorderThin = 8388608,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLBorder_Max = 2147483647
	}
}