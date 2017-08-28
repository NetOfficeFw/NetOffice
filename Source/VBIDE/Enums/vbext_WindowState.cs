using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsEnum)]
	public enum vbext_WindowState
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_ws_Normal = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_ws_Minimize = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_ws_Maximize = 2
	}
}