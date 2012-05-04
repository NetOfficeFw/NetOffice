using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum RecordsetOptionEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDenyWrite = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDenyRead = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbReadOnly = 4,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbAppendOnly = 8,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbInconsistent = 16,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbConsistent = 32,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSQLPassThrough = 64,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbFailOnError = 128,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbForwardOnly = 256,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSeeChanges = 512,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbRunAsync = 1024,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbExecDirect = 2048
	}
}