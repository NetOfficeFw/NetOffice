using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 3.6, 12.0
	 /// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsEnum)]
	public enum PermissionEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecNoAccess = 0,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1048575</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecFullAccess = 1048575,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDelete = 65536,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecReadSec = 131072,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecWriteSec = 262144,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecWriteOwner = 524288,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDBCreate = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDBOpen = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDBExclusive = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDBAdmin = 8,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecCreate = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecReadDef = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>65548</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecWriteDef = 65548,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecRetrieveData = 20,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecInsertData = 32,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecReplaceData = 64,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSecDeleteData = 128
	}
}