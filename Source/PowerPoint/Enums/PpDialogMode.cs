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
	public enum PpDialogMode
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogModeMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogModeless = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppDialogModal = 1
	}
}