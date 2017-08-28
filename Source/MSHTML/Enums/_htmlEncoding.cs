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
	public enum _htmlEncoding
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlEncodingURL = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlEncodingMultipart = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlEncodingText = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlEncoding_Max = 2147483647
	}
}