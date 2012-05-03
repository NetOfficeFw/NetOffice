using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlCopyPictureFormat
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlBitmap = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>-4147</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlPicture = -4147
	}
}