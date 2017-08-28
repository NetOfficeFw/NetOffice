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
	public enum CommitTransOptionsEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbForceOSFlush = 1
	}
}