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
	public enum XlBinsType
	{
		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlBinsTypeAutomatic = 0,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlBinsTypeCategorical = 1,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlBinsTypeManual = 2,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlBinsTypeBinSize = 3,

		 /// <summary>
		 /// SupportByVersion Excel 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 16)]
		 xlBinsTypeBinCount = 4
	}
}