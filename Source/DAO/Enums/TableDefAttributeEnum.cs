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
	public enum TableDefAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbAttachExclusive = 65536,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbAttachSavePWD = 131072,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>-2147483646</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbSystemObject = -2147483646,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1073741824</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbAttachedTable = 1073741824,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>536870912</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbAttachedODBC = 536870912,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbHiddenObject = 1
	}
}