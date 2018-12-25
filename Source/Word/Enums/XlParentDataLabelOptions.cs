using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 16
	 /// </summary>
	[SupportByVersion("Word", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlParentDataLabelOptions
	{
		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 16)]
		 xlParentDataLabelOptionsNone = 0,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 16)]
		 xlParentDataLabelOptionsBanner = 1,

		 /// <summary>
		 /// SupportByVersion Word 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 16)]
		 xlParentDataLabelOptionsOverlapping = 2
	}
}