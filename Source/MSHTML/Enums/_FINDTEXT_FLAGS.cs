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
	public enum _FINDTEXT_FLAGS
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_BACKWARDS = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_WHOLEWORD = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_MATCHCASE = 4,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_RAW = 131072,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>536870912</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_MATCHDIAC = 536870912,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1073741824</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_MATCHKASHIDA = 1073741824,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_MATCHALEFHAMZA = -2147483648,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 FINDTEXT_FLAGS_Max = 2147483647
	}
}