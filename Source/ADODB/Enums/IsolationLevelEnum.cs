using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum IsolationLevelEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactUnspecified = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactChaos = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactReadUncommitted = 256,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactBrowse = 256,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactCursorStability = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactReadCommitted = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactRepeatableRead = 65536,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactSerializable = 1048576,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adXactIsolated = 1048576
	}
}