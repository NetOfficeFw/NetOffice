using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msochartfieldtype-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoChartFieldType
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldBubbleSize = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldCategoryName = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldPercentage = 3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldSeriesName = 4,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldValue = 5,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoChartFieldFormula = 6
	}
}