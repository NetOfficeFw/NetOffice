using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum LockTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbPessimistic = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbOptimistic = 3,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbOptimisticValue = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbOptimisticBatch = 5
	}
}