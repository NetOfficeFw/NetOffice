using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 14, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867005(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjFieldTypes
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjCostField = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjDateField = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjDurationField = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjFlagField = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjFinishField = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjNumberField = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjStartField = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTextField = 7
	}
}