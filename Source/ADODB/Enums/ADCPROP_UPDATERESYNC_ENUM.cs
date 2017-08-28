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
	public enum ADCPROP_UPDATERESYNC_ENUM
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncNone = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncAutoIncrement = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncConflicts = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncUpdates = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncInserts = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adResyncAll = 15
	}
}