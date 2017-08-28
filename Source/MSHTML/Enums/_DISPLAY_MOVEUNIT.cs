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
	public enum _DISPLAY_MOVEUNIT
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_PreviousLine = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_NextLine = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_CurrentLineStart = 3,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_CurrentLineEnd = 4,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_TopOfWindow = 5,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_BottomOfWindow = 6,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 DISPLAY_MOVEUNIT_Max = 2147483647
	}
}