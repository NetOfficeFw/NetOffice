using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XactAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactCommitRetaining = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactAbortRetaining = 262144,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactAsyncPhaseOne = 524288,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactSyncPhaseOne = 1048576
	}
}