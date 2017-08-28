using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861393(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjCompareVersionItems
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsAllDifferences = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsChangedItems = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsUnchangedItems = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsCommonItems = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsUniqueItemsOfVersion1 = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsUniqueItemsOfVersion2 = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,14)]
		 pjCompareVersionItemsAllItems = 6
	}
}