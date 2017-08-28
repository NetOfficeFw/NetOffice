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
	public enum _BEHAVIOR_RENDER_INFO
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_BEFOREBACKGROUND = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_AFTERBACKGROUND = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_BEFORECONTENT = 4,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_AFTERCONTENT = 8,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_AFTERFOREGROUND = 32,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_ABOVECONTENT = 40,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>255</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_ALLLAYERS = 255,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_DISABLEBACKGROUND = 256,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_DISABLENEGATIVEZ = 512,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_DISABLECONTENT = 1024,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_DISABLEPOSITIVEZ = 2048,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>3840</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_DISABLEALLLAYERS = 3840,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_HITTESTING = 4096,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_SURFACE = 1048576,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2097152</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORRENDERINFO_3DSURFACE = 2097152,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIOR_RENDER_INFO_Max = 2147483647
	}
}