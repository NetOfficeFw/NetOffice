using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlEndStyleCap
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlCap = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 xlNoCap = 2
	}
}