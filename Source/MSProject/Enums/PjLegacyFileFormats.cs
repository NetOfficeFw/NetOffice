using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjLegacyFileFormats
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjDoNotLoadLegacyFiles = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjPromptBeforeLoadingLegacyFiles = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjEnableLegacyFilesLoad = 2
	}
}