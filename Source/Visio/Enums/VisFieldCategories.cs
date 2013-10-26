using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisFieldCategories
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatCustom = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatDateTime = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatDocument = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatGeometry = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatObject = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatPage = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visFCatNotes = 6
	}
}