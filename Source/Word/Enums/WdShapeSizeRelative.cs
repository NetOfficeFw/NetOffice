using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdShapeSizeRelative"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdShapeSizeRelative
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-999999</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdShapeSizeRelativeNone = -999999
	}
}