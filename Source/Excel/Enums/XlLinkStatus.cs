using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlLinkStatus"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlLinkStatus
	{
		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusOK = 0,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusMissingFile = 1,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusMissingSheet = 2,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusOld = 3,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusSourceNotCalculated = 4,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusIndeterminate = 5,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusNotStarted = 6,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusInvalidName = 7,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusSourceNotOpen = 8,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusSourceOpen = 9,

		 /// <summary>
		 /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Excel", 10,11,12,14,15,16)]
		 xlLinkStatusCopiedValues = 10
	}
}