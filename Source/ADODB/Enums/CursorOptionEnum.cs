using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum CursorOptionEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adHoldRecords = 256,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adMovePrevious = 512,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16778240</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adAddNew = 16778240,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16779264</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adDelete = 16779264,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16809984</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adUpdate = 16809984,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adBookmark = 8192,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adApproxPosition = 16384,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adUpdateBatch = 65536,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adResync = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adNotify = 262144,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFind = 524288,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4194304</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adSeek = 4194304,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8388608</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adIndex = 8388608
	}
}