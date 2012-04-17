using System;
using LateBindingApi.Core;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoSmartArtNodePosition
	{
		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeDefault = 1,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeAfter = 2,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeBefore = 3,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeAbove = 4,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSmartArtNodeBelow = 5
	}
}