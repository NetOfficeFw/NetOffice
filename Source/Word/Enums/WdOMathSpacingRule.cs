using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdOMathSpacingRule"/> </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdOMathSpacingRule
	{
		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdOMathSpacingSingle = 0,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdOMathSpacing1pt5 = 1,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdOMathSpacingDouble = 2,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdOMathSpacingExactly = 3,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdOMathSpacingMultiple = 4
	}
}