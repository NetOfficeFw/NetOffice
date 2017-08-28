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
	public enum XactAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adXactCommitRetaining = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adXactAbortRetaining = 262144,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adXactAsyncPhaseOne = 524288,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adXactSyncPhaseOne = 1048576
	}
}