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
	public enum _htmlMarqueeDirection
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlMarqueeDirectionleft = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlMarqueeDirectionright = 3,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlMarqueeDirectionup = 5,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlMarqueeDirectiondown = 7,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 htmlMarqueeDirection_Max = 2147483647
	}
}