using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdEditorType"/> </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdEditorType
	{
		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdEditorEveryone = -1,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdEditorOwners = -4,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdEditorEditors = -5,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdEditorCurrent = -6
	}
}