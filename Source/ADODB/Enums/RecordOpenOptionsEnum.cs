using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum RecordOpenOptionsEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adOpenRecordUnspecified = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8388608</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adOpenSource = 8388608,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adOpenAsync = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adDelayFetchStream = 16384,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adDelayFetchFields = 32768
	}
}