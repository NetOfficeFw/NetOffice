using System;
using LateBindingApi.Core;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Excel", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlListDataType
	{
		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeNone = 0,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeText = 1,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeMultiLineText = 2,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeNumber = 3,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeCurrency = 4,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeDateTime = 5,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeChoice = 6,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeChoiceMulti = 7,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeListLookup = 8,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeCheckbox = 9,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeHyperLink = 10,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeCounter = 11,

		 /// <summary>
		 /// SupportByVersion Excel 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Excel", 11,12,14)]
		 xlListDataTypeMultiLineRichText = 12
	}
}