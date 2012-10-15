using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjFieldTypes
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjCostField = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjDateField = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjDurationField = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjFlagField = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjFinishField = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjNumberField = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjStartField = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjTextField = 7
	}
}