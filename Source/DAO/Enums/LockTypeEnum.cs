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
	public enum LockTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbPessimistic = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbOptimistic = 3,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbOptimisticValue = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbOptimisticBatch = 5
	}
}