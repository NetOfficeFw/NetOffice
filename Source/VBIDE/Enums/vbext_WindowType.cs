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
	public enum vbext_WindowType
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_CodeWindow = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Designer = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Browser = 2,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Watch = 3,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Locals = 4,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Immediate = 5,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_ProjectWindow = 6,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_PropertyWindow = 7,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Find = 8,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_FindReplace = 9,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_Toolbox = 10,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_LinkedWindowFrame = 11,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_MainWindow = 12,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_wt_ToolWindow = 15
	}
}