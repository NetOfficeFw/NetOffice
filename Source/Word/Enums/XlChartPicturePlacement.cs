using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.xlchartpictureplacement"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlChartPicturePlacement
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlSides = 1,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlEnd = 2,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlEndSides = 3,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlFront = 4,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlFrontSides = 5,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlFrontEnd = 6,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlAllFaces = 7
	}
}