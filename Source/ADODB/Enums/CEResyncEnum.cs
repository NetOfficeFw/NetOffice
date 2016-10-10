using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum CEResyncEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncNone = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncAutoIncrement = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncConflicts = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncUpdates = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncInserts = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1)]
		 adResyncAll = 15
	}
}