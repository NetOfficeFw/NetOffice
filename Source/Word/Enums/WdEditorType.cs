using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdEditorType
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdEditorEveryone = -1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdEditorOwners = -4,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdEditorEditors = -5,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdEditorCurrent = -6
	}
}