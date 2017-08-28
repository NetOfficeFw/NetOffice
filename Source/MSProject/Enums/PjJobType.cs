using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11
	 /// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsEnum)]
	public enum PjJobType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheProjectSave = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheProjectCheckin = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheProjectPublish = 2
	}
}