using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767414(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("Visio", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisResizeDirection
	{
		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirE = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirNE = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirN = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirNW = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirW = 4,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirSW = 5,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirS = 6,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visResizeDirSE = 7
	}
}