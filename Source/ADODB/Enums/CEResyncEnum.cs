using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1
	 /// </summary>
	[SupportByVersion("ADODB", 2.1)]
	[EntityType(EntityType.IsEnum)]
	public enum CEResyncEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncNone = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncAutoIncrement = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncConflicts = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncUpdates = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncInserts = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adResyncAll = 15
	}
}