using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 16
	 /// </summary>
	[SupportByVersion("Word", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlBinsType
	{
		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 16)]
		 xlBinsTypeAutomatic = 0,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 16)]
		 xlBinsTypeCategorical = 1,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 16)]
		 xlBinsTypeManual = 2,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 16)]
		 xlBinsTypeBinSize = 3,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 16)]
		 xlBinsTypeBinCount = 4
	}
}