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
	public enum _HTMLBorderStyle
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStyleNormal = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStyleRaised = 256,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStyleSunken = 512,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>768</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStylecombined = 768,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStyleStatic = 131072,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 HTMLBorderStyle_Max = 2147483647
	}
}