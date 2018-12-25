using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 16
	 /// </summary>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlBinsType
	{
		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 16)]
		 xlBinsTypeAutomatic = 0,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 16)]
		 xlBinsTypeCategorical = 1,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 16)]
		 xlBinsTypeManual = 2,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 16)]
		 xlBinsTypeBinSize = 3,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 16)]
		 xlBinsTypeBinCount = 4
	}
}