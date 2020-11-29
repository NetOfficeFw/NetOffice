using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimAfterEffect"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimAfterEffect
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimAfterEffectMixed = -1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimAfterEffectNone = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimAfterEffectDim = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimAfterEffectHide = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimAfterEffectHideOnNextClick = 3
	}
}