using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjPoolOpen
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPromptPool = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPoolReadOnly = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPoolReadWrite = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPoolAndSharers = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjDoNotOpenPool = 4
	}
}