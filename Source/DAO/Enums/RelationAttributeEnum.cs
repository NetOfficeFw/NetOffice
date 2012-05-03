using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 12, 3.6
	 /// </summary>
	[SupportByVersionAttribute("DAO", 12,3.6)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum RelationAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationUnique = 1,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationDontEnforce = 2,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationInherited = 4,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationUpdateCascade = 256,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationDeleteCascade = 4096,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>16777216</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationLeft = 16777216,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>33554432</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbRelationRight = 33554432
	}
}