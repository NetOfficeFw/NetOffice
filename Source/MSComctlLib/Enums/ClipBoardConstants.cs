using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSComctlLibApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSComctlLib 6
	 /// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsEnum)]
	public enum ClipBoardConstants
	{
		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFText = 1,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFBitmap = 2,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFMetafile = 3,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFDIB = 8,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFPalette = 9,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFEMetafile = 14,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFFiles = 15,

		 /// <summary>
		 /// SupportByVersion MSComctlLib 6
		 /// </summary>
		 /// <remarks>-16639</remarks>
		 [SupportByVersion("MSComctlLib", 6)]
		 ccCFRTF = -16639
	}
}