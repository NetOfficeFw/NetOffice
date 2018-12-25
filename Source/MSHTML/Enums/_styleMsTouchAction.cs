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
	public enum _styleMsTouchAction
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionNotSet = -1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionNone = 0,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionAuto = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionManipulation = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionDoubleTapZoom = 4,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionPanX = 8,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionPanY = 16,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionPinchZoom = 32,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionCrossSlideX = 64,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchActionCrossSlideY = 128,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 styleMsTouchAction_Max = 2147483647
	}
}