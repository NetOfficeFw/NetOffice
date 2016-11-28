using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum RecordStatusEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecOK = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecNew = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecModified = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecDeleted = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecUnmodified = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecInvalid = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecMultipleChanges = 64,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecPendingChanges = 128,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecCanceled = 256,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecCantRelease = 1024,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecConcurrencyViolation = 2048,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecIntegrityViolation = 4096,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecMaxChangesExceeded = 8192,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecObjectOpen = 16384,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecOutOfMemory = 32768,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecPermissionDenied = 65536,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecSchemaViolation = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRecDBDeleted = 262144
	}
}