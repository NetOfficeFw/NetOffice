using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 5, 12
	 /// </summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum FieldAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbFixedField = 1,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbVariableField = 2,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbAutoIncrField = 16,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbUpdatableField = 32,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbSystemField = 8192,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbHyperlinkField = 32768,

		 /// <summary>
		 /// SupportByVersion DAO 5, 12
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 5,12)]
		 dbDescending = 1
	}
}