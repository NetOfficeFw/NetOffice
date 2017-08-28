using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoDiagramNodeType
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoDiagramNode = 1,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoDiagramAssistant = 2
	}
}