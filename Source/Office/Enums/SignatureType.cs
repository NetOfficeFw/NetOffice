using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum SignatureType
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 sigtypeUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 sigtypeNonVisible = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 sigtypeSignatureLine = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 sigtypeMax = 3
	}
}