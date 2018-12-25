using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 16
	 /// </summary>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlParentDataLabelOptions
	{
		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 16)]
		 xlParentDataLabelOptionsNone = 0,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 16)]
		 xlParentDataLabelOptionsBanner = 1,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 16)]
		 xlParentDataLabelOptionsOverlapping = 2
	}
}