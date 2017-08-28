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
	public enum PositionEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPosUnknown = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPosBOF = -2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPosEOF = -3
	}
}