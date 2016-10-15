using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229719.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCategoryLabelLevel
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 xlCategoryLabelLevelNone = -3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 xlCategoryLabelLevelCustom = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 xlCategoryLabelLevelAll = -1
	}
}