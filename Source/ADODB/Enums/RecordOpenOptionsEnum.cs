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
	public enum RecordOpenOptionsEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adOpenRecordUnspecified = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8388608</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adOpenSource = 8388608,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adOpenAsync = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adDelayFetchStream = 16384,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adDelayFetchFields = 32768
	}
}