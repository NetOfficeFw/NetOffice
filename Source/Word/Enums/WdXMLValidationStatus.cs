using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdXMLValidationStatus
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLValidationStatusOK = 0,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1072898048</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdXMLValidationStatusCustom = -1072898048
	}
}