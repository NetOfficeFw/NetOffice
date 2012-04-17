using System;
using LateBindingApi.Core;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 12, 3.6
	 /// </summary>
	[SupportByVersionAttribute("DAO", 12,3.6)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum FieldAttributeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbFixedField = 1,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbVariableField = 2,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbAutoIncrField = 16,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbUpdatableField = 32,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSystemField = 8192,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbHyperlinkField = 32768,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbDescending = 1
	}
}