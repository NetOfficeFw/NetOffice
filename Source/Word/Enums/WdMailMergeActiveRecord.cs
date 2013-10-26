using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196015.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdMailMergeActiveRecord
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoActiveRecord = -1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNextRecord = -2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdPreviousRecord = -3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFirstRecord = -4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLastRecord = -5,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdFirstDataSourceRecord = -6,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdLastDataSourceRecord = -7,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdNextDataSourceRecord = -8,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdPreviousDataSourceRecord = -9
	}
}