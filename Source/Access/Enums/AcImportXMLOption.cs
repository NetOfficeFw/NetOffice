using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcImportXMLOption
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acStructureOnly = 0,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acStructureAndData = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14,15)]
		 acAppendData = 2
	}
}