using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9
	 /// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsEnum)]
	public enum PpDialogStyle
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogStyleMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogStandard = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogTabbed = 2
	}
}