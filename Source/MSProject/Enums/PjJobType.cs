using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 16
	 /// </summary>
	[SupportByVersion("MSProject", 11,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjJobType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheProjectSave = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheProjectCheckin = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheProjectPublish = 2
	}
}