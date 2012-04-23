using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSmartArtNodeType
	{
		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeTypeDefault = 1,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeTypeAssistant = 2
	}
}