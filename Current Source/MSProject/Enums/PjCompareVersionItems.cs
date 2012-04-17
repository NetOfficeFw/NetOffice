using System;
using LateBindingApi.Core;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjCompareVersionItems
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsAllDifferences = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsChangedItems = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsUnchangedItems = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsCommonItems = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsUniqueItemsOfVersion1 = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsUniqueItemsOfVersion2 = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjCompareVersionItemsAllItems = 6
	}
}