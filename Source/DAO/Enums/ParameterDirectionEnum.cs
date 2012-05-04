using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum ParameterDirectionEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbParamInput = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbParamOutput = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbParamInputOutput = 3,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbParamReturnValue = 4
	}
}