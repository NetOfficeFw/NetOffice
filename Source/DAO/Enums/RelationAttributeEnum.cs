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
	public enum RelationAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationUnique = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationDontEnforce = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationInherited = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationUpdateCascade = 256,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationDeleteCascade = 4096,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>16777216</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationLeft = 16777216,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>33554432</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbRelationRight = 33554432
	}
}