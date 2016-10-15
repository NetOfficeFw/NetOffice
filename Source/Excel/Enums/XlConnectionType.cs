using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839396.aspx </remarks>
	[SupportByVersionAttribute("Excel", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlConnectionType
	{
		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlConnectionTypeOLEDB = 1,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlConnectionTypeODBC = 2,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlConnectionTypeXMLMAP = 3,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlConnectionTypeTEXT = 4,

		 /// <summary>
		 /// SupportByVersion Excel 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 12,14,15,16)]
		 xlConnectionTypeWEB = 5,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlConnectionTypeDATAFEED = 6,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlConnectionTypeMODEL = 7,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlConnectionTypeWORKSHEET = 8,

		 /// <summary>
		 /// SupportByVersion Excel 15,16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 15, 16)]
		 xlConnectionTypeNOSOURCE = 9
	}
}