using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum DriverPromptEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDriverPrompt = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDriverNoPrompt = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDriverComplete = 0,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDriverCompleteRequired = 3
	}
}