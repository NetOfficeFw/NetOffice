using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230235.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpGuideOrientation
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppHorizontalGuide = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 15, 16)]
		 ppVerticalGuide = 2
	}
}