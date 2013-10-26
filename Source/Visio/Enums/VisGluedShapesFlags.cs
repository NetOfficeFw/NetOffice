using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768275(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("Visio", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisGluedShapesFlags
	{
		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesAll1D = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesIncoming1D = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesOutgoing1D = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesAll2D = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesIncoming2D = 4,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visGluedShapesOutgoing2D = 5
	}
}