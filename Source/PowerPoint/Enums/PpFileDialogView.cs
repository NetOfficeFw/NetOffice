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
	public enum PpFileDialogView
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppFileDialogViewDetails = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppFileDialogViewPreview = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppFileDialogViewProperties = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 9)]
		 ppFileDialogViewList = 4
	}
}