using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum StreamTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adTypeBinary = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adTypeText = 2
	}
}