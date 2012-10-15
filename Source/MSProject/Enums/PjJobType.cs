using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjJobType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheProjectSave = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheProjectCheckin = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheProjectPublish = 2
	}
}