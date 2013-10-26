using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835490.aspx </remarks>
	[SupportByVersionAttribute("Word", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdWrapTypeMerged
	{
		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeInline = 0,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeSquare = 1,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeTight = 2,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeBehind = 3,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeFront = 4,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeThrough = 5,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdWrapMergeTopBottom = 6
	}
}