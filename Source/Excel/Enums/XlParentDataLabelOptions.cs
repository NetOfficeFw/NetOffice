using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 16
	 /// </summary>
	[SupportByVersion("Excel", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlParentDataLabelOptions
	{
		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlParentDataLabelOptionsNone = 0,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlParentDataLabelOptionsBanner = 1,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlParentDataLabelOptionsOverlapping = 2
	}
}