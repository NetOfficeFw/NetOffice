using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863548(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjMonthLabel
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mm = 57,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mm_yy = 86,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mm_yyy = 85,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_m = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mmm = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mmm_yyy = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mmmm = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonth_mmmm_yyyy = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromEnd_mm = 59,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromEnd_Mmm = 58,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromEnd_Month_mm = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromStart_mm = 61,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromStart_Mmm = 60,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelMonthFromStart_Month_mm = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthLabelNoDateFormat = 35
	}
}