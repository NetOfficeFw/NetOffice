using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisKeyButtonFlags
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visMouseLeft = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visMouseMiddle = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visMouseRight = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visKeyShift = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visKeyControl = 8
	}
}