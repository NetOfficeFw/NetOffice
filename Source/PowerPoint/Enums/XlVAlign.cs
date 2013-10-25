using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746626.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlVAlign
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4107</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlVAlignBottom = -4107,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4108</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlVAlignCenter = -4108,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4117</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlVAlignDistributed = -4117,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4130</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlVAlignJustify = -4130,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15
		 /// </summary>
		 /// <remarks>-4160</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14,15)]
		 xlVAlignTop = -4160
	}
}