using System;
using NetOffice;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum _HTMLDlgBorder
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLDlgBorderThin = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLDlgBorderThick = 262144,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 HTMLDlgBorder_Max = 2147483647
	}
}