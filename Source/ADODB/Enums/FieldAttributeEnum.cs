using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum FieldAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldUnspecified = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldMayDefer = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldUpdatable = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldUnknownUpdatable = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldFixed = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldIsNullable = 32,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldMayBeNull = 64,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldLong = 128,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldRowID = 256,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldRowVersion = 512,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldCacheDeferred = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldNegativeScale = 16384,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adFldKeyColumn = 32768,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFldIsChapter = 8192,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFldIsRowURL = 65536,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFldIsDefaultStream = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFldIsCollection = 262144
	}
}