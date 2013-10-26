using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisDataColumnProperties
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyType = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyLangID = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyCalendar = 3,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyUnits = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyCurrency = 5,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyDisplayName = 6,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyVisible = 7,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataColumnPropertyHyperlink = 8
	}
}