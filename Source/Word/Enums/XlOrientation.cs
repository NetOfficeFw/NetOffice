using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.xlorientation"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlOrientation
	{
		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4170</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlDownward = -4170,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4128</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlHorizontal = -4128,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4171</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlUpward = -4171,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>-4166</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 xlVertical = -4166
	}
}