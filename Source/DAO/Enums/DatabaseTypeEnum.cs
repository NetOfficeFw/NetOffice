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
	public enum DatabaseTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbVersion10 = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbEncrypt = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbDecrypt = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbVersion11 = 8,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbVersion20 = 16,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbVersion30 = 32,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbVersion40 = 64,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("DAO", 12.0)]
		 dbVersion120 = 128,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("DAO", 12.0)]
		 dbVersion140 = 256,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("DAO", 12.0)]
		 dbVersion150 = 512
	}
}