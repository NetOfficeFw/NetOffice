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
	public enum UpdateCriteriaEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaKey = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaModValues = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaAllCols = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaTimestamp = 8,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaDeleteInsert = 16,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbCriteriaUpdate = 32
	}
}