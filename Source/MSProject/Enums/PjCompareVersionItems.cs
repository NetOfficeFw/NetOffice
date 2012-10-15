using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjCompareVersionItems
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsAllDifferences = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsChangedItems = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsUnchangedItems = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsCommonItems = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsUniqueItemsOfVersion1 = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsUniqueItemsOfVersion2 = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCompareVersionItemsAllItems = 6
	}
}