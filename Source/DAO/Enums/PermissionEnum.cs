using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PermissionEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecNoAccess = 0,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1048575</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecFullAccess = 1048575,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDelete = 65536,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecReadSec = 131072,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecWriteSec = 262144,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecWriteOwner = 524288,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDBCreate = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDBOpen = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDBExclusive = 4,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDBAdmin = 8,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecCreate = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecReadDef = 4,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>65548</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecWriteDef = 65548,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecRetrieveData = 20,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecInsertData = 32,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecReplaceData = 64,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSecDeleteData = 128
	}
}