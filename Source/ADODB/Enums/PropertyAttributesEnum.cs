using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum PropertyAttributesEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropNotSupported = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropRequired = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropOptional = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropRead = 512,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropWrite = 1024
	}
}