using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 12, 3.6
	 /// </summary>
	[SupportByVersionAttribute("DAO", 12,3.6)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum LockTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbPessimistic = 2,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbOptimistic = 3,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbOptimisticValue = 1,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbOptimisticBatch = 5
	}
}