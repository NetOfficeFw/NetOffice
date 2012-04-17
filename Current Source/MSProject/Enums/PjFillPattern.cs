using System;
using LateBindingApi.Core;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjFillPattern
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjHollowPattern = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjSolidFillPattern = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjLightFillPattern = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjMediumFillPattern = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjDarkFillPattern = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjDiagonalLeftPattern = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjDiagonalRightPattern = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjDiagonalCrossPattern = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjLineVerticalPattern = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjLineHorizontalPattern = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjLineCrossPattern = 10
	}
}