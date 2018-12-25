using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867782(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjField
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743680</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWork = 188743680,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743681</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineWork = 188743681,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743682</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualWork = 188743682,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743683</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWorkVariance = 188743683,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743684</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRemainingWork = 188743684,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743685</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost = 188743685,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743686</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineCost = 188743686,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743687</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualCost = 188743687,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743688</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFixedCost = 188743688,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743689</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCostVariance = 188743689,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743690</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRemainingCost = 188743690,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743691</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBCWP = 188743691,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743692</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBCWS = 188743692,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743693</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSV = 188743693,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743694</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskName = 188743694,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743695</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNotes = 188743695,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743696</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWBS = 188743696,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743697</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskConstraintType = 188743697,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743698</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskConstraintDate = 188743698,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743699</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCritical = 188743699,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743700</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLevelDelay = 188743700,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743701</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFreeSlack = 188743701,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743702</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskTotalSlack = 188743702,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743703</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskID = 188743703,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743704</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskMilestone = 188743704,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743705</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPriority = 188743705,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743706</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSubproject = 188743706,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743707</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineDuration = 188743707,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743708</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualDuration = 188743708,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743709</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration = 188743709,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743710</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDurationVariance = 188743710,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743711</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRemainingDuration = 188743711,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743712</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPercentComplete = 188743712,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743713</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPercentWorkComplete = 188743713,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743714</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFixedDuration = 188743714,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743715</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart = 188743715,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743716</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish = 188743716,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743717</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEarlyStart = 188743717,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743718</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEarlyFinish = 188743718,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743719</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLateStart = 188743719,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743720</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLateFinish = 188743720,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743721</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualStart = 188743721,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743722</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualFinish = 188743722,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743723</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineStart = 188743723,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743724</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineFinish = 188743724,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743725</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStartVariance = 188743725,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743726</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinishVariance = 188743726,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743727</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPredecessors = 188743727,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743728</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSuccessors = 188743728,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743729</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceNames = 188743729,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743730</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceInitials = 188743730,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743731</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText1 = 188743731,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743732</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart1 = 188743732,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743733</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish1 = 188743733,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743734</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText2 = 188743734,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743735</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart2 = 188743735,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743736</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish2 = 188743736,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743737</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText3 = 188743737,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743738</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart3 = 188743738,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743739</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish3 = 188743739,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743740</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText4 = 188743740,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743741</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart4 = 188743741,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743742</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish4 = 188743742,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743743</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText5 = 188743743,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743744</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart5 = 188743744,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743745</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish5 = 188743745,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743746</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText6 = 188743746,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743747</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText7 = 188743747,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743748</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText8 = 188743748,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743749</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText9 = 188743749,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743750</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText10 = 188743750,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743751</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskMarked = 188743751,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743752</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag1 = 188743752,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743753</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag2 = 188743753,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743754</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag3 = 188743754,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743755</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag4 = 188743755,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743756</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag5 = 188743756,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743757</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag6 = 188743757,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743758</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag7 = 188743758,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743759</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag8 = 188743759,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743760</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag9 = 188743760,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743761</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag10 = 188743761,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743762</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRollup = 188743762,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743763</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCV = 188743763,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743764</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskProject = 188743764,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743765</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineLevel = 188743765,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743766</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskUniqueID = 188743766,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743767</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber1 = 188743767,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743768</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber2 = 188743768,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743769</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber3 = 188743769,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743770</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber4 = 188743770,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743771</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber5 = 188743771,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743772</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSummary = 188743772,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743773</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCreated = 188743773,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743774</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSheetNotes = 188743774,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743775</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskUniquePredecessors = 188743775,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743776</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskUniqueSuccessors = 188743776,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743777</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskObjects = 188743777,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743778</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLinkedFields = 188743778,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743779</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResume = 188743779,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743780</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStop = 188743780,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743781</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResumeNoEarlierThan = 188743781,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743782</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineNumber = 188743782,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743783</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration1 = 188743783,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743784</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration2 = 188743784,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743785</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration3 = 188743785,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743786</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost1 = 188743786,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743787</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost2 = 188743787,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743788</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost3 = 188743788,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743789</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHideBar = 188743789,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743790</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskConfirmed = 188743790,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743791</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskUpdateNeeded = 188743791,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743792</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskContact = 188743792,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743793</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceGroup = 188743793,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743800</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskACWP = 188743800,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743808</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskType = 188743808,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743809</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRecurring = 188743809,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743812</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEffortDriven = 188743812,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743815</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskParentTask = 188743815,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743843</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOvertimeWork = 188743843,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743844</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualOvertimeWork = 188743844,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743845</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRemainingOvertimeWork = 188743845,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743846</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRegularWork = 188743846,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743848</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOvertimeCost = 188743848,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743849</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualOvertimeCost = 188743849,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743850</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRemainingOvertimeCost = 188743850,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743880</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFixedCostAccrual = 188743880,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743885</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskIndicators = 188743885,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743897</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHyperlink = 188743897,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743898</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHyperlinkAddress = 188743898,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743899</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHyperlinkSubAddress = 188743899,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743900</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHyperlinkHref = 188743900,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743904</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskIsAssignment = 188743904,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743905</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOverallocated = 188743905,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743912</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskExternalTask = 188743912,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743926</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSubprojectReadOnly = 188743926,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743930</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResponsePending = 188743930,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743931</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskTeamStatusPending = 188743931,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743932</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLevelCanSplit = 188743932,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743933</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskLevelAssignments = 188743933,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743936</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWorkContour = 188743936,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743938</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost4 = 188743938,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743939</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost5 = 188743939,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743940</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost6 = 188743940,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743941</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost7 = 188743941,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743942</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost8 = 188743942,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743943</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost9 = 188743943,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743944</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCost10 = 188743944,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743945</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate1 = 188743945,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743946</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate2 = 188743946,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743947</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate3 = 188743947,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743948</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate4 = 188743948,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743949</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate5 = 188743949,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743950</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate6 = 188743950,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743951</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate7 = 188743951,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743952</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate8 = 188743952,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743953</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate9 = 188743953,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743954</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDate10 = 188743954,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743955</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration4 = 188743955,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743956</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration5 = 188743956,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743957</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration6 = 188743957,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743958</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration7 = 188743958,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743959</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration8 = 188743959,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743960</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration9 = 188743960,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743961</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration10 = 188743961,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743962</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart6 = 188743962,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743963</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish6 = 188743963,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743964</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart7 = 188743964,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743965</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish7 = 188743965,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743966</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart8 = 188743966,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743967</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish8 = 188743967,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743968</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart9 = 188743968,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743969</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish9 = 188743969,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743970</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStart10 = 188743970,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743971</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinish10 = 188743971,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743972</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag11 = 188743972,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743973</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag12 = 188743973,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743974</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag13 = 188743974,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743975</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag14 = 188743975,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743976</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag15 = 188743976,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743977</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag16 = 188743977,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743978</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag17 = 188743978,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743979</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag18 = 188743979,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743980</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag19 = 188743980,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743981</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFlag20 = 188743981,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743982</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber6 = 188743982,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743983</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber7 = 188743983,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743984</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber8 = 188743984,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743985</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber9 = 188743985,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743986</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber10 = 188743986,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743987</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber11 = 188743987,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743988</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber12 = 188743988,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743989</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber13 = 188743989,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743990</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber14 = 188743990,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743991</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber15 = 188743991,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743992</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber16 = 188743992,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743993</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber17 = 188743993,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743994</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber18 = 188743994,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743995</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber19 = 188743995,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743996</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskNumber20 = 188743996,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743997</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText11 = 188743997,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743998</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText12 = 188743998,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743999</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText13 = 188743999,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744000</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText14 = 188744000,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744001</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText15 = 188744001,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744002</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText16 = 188744002,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744003</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText17 = 188744003,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744004</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText18 = 188744004,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744005</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText19 = 188744005,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744006</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText20 = 188744006,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744007</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText21 = 188744007,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744008</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText22 = 188744008,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744009</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText23 = 188744009,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744010</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText24 = 188744010,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744011</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText25 = 188744011,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744012</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText26 = 188744012,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744013</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText27 = 188744013,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744014</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText28 = 188744014,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744015</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText29 = 188744015,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744016</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskText30 = 188744016,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744029</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourcePhonetics = 188744029,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744040</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskIndex = 188744040,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744046</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskAssignmentDelay = 188744046,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744047</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskAssignmentUnits = 188744047,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744048</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCostRateTable = 188744048,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744049</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPreleveledStart = 188744049,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744050</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPreleveledFinish = 188744050,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744076</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEstimated = 188744076,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744079</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskIgnoreResourceCalendar = 188744079,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744082</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCalendar = 188744082,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744083</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration1Estimated = 188744083,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744084</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration2Estimated = 188744084,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744085</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration3Estimated = 188744085,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744086</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration4Estimated = 188744086,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744087</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration5Estimated = 188744087,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744088</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration6Estimated = 188744088,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744089</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration7Estimated = 188744089,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744090</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration8Estimated = 188744090,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744091</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration9Estimated = 188744091,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744092</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDuration10Estimated = 188744092,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744093</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineDurationEstimated = 188744093,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744096</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode1 = 188744096,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744098</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode2 = 188744098,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744100</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode3 = 188744100,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744102</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode4 = 188744102,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744104</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode5 = 188744104,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744106</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode6 = 188744106,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744108</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode7 = 188744108,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744110</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode8 = 188744110,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744112</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode9 = 188744112,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744114</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskOutlineCode10 = 188744114,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744117</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeadline = 188744117,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744118</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStartSlack = 188744118,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744119</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskFinishSlack = 188744119,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744121</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskVAC = 188744121,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744126</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskGroupBySummary = 188744126,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744129</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWBSPredecessors = 188744129,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744130</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskWBSSuccessors = 188744130,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744131</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceType = 188744131,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744132</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskHyperlinkScreenTip = 188744132,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744162</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1Start = 188744162,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744163</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1Finish = 188744163,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744164</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1Cost = 188744164,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744165</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1Work = 188744165,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744167</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1Duration = 188744167,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744173</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2Start = 188744173,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744174</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2Finish = 188744174,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744175</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2Cost = 188744175,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744176</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2Work = 188744176,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744178</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2Duration = 188744178,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744184</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3Start = 188744184,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744185</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3Finish = 188744185,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744186</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3Cost = 188744186,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744187</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3Work = 188744187,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744189</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3Duration = 188744189,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744195</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4Start = 188744195,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744196</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4Finish = 188744196,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744197</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4Cost = 188744197,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744198</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4Work = 188744198,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744200</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4Duration = 188744200,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744206</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5Start = 188744206,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744207</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5Finish = 188744207,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744208</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5Cost = 188744208,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744209</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5Work = 188744209,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744211</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5Duration = 188744211,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744217</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCPI = 188744217,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744218</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSPI = 188744218,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744219</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCVPercent = 188744219,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744220</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskSVPercent = 188744220,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744221</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEAC = 188744221,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744222</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskTCPI = 188744222,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744223</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStatus = 188744223,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744224</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6Start = 188744224,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744225</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6Finish = 188744225,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744226</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6Cost = 188744226,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744227</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6Work = 188744227,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744229</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6Duration = 188744229,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744235</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7Start = 188744235,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744236</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7Finish = 188744236,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744237</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7Cost = 188744237,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744238</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7Work = 188744238,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744240</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7Duration = 188744240,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744246</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8Start = 188744246,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744247</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8Finish = 188744247,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744248</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8Cost = 188744248,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744249</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8Work = 188744249,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744251</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8Duration = 188744251,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744257</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9Start = 188744257,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744258</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9Finish = 188744258,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744259</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9Cost = 188744259,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744260</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9Work = 188744260,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744262</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9Duration = 188744262,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744268</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10Start = 188744268,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744269</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10Finish = 188744269,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744270</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10Cost = 188744270,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744271</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10Work = 188744271,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744273</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10Duration = 188744273,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744279</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost1 = 188744279,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744280</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost2 = 188744280,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744281</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost3 = 188744281,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744282</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost4 = 188744282,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744283</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost5 = 188744283,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744284</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost6 = 188744284,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744285</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost7 = 188744285,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744286</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost8 = 188744286,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744287</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost9 = 188744287,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744288</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseCost10 = 188744288,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744289</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate1 = 188744289,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744290</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate2 = 188744290,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744291</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate3 = 188744291,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744292</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate4 = 188744292,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744293</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate5 = 188744293,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744294</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate6 = 188744294,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744295</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate7 = 188744295,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744296</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate8 = 188744296,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744297</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate9 = 188744297,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744298</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate10 = 188744298,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744299</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate11 = 188744299,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744300</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate12 = 188744300,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744301</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate13 = 188744301,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744302</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate14 = 188744302,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744303</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate15 = 188744303,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744304</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate16 = 188744304,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744305</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate17 = 188744305,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744306</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate18 = 188744306,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744307</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate19 = 188744307,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744308</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate20 = 188744308,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744309</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate21 = 188744309,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744310</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate22 = 188744310,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744311</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate23 = 188744311,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744312</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate24 = 188744312,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744313</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate25 = 188744313,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744314</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate26 = 188744314,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744315</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate27 = 188744315,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744316</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate28 = 188744316,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744317</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate29 = 188744317,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744318</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDate30 = 188744318,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744319</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration1 = 188744319,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744320</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration2 = 188744320,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744321</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration3 = 188744321,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744322</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration4 = 188744322,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744323</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration5 = 188744323,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744324</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration6 = 188744324,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744325</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration7 = 188744325,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744326</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration8 = 188744326,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744327</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration9 = 188744327,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744328</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseDuration10 = 188744328,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744339</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag1 = 188744339,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744340</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag2 = 188744340,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744341</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag3 = 188744341,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744342</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag4 = 188744342,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744343</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag5 = 188744343,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744344</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag6 = 188744344,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744345</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag7 = 188744345,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744346</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag8 = 188744346,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744347</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag9 = 188744347,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744348</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag10 = 188744348,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744349</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag11 = 188744349,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744350</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag12 = 188744350,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744351</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag13 = 188744351,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744352</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag14 = 188744352,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744353</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag15 = 188744353,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744354</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag16 = 188744354,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744355</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag17 = 188744355,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744356</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag18 = 188744356,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744357</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag19 = 188744357,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744358</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseFlag20 = 188744358,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744379</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber1 = 188744379,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744380</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber2 = 188744380,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744381</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber3 = 188744381,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744382</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber4 = 188744382,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744383</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber5 = 188744383,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744384</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber6 = 188744384,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744385</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber7 = 188744385,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744386</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber8 = 188744386,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744387</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber9 = 188744387,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744388</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber10 = 188744388,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744389</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber11 = 188744389,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744390</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber12 = 188744390,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744391</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber13 = 188744391,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744392</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber14 = 188744392,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744393</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber15 = 188744393,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744394</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber16 = 188744394,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744395</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber17 = 188744395,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744396</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber18 = 188744396,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744397</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber19 = 188744397,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744398</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber20 = 188744398,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744399</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber21 = 188744399,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744400</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber22 = 188744400,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744401</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber23 = 188744401,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744402</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber24 = 188744402,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744403</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber25 = 188744403,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744404</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber26 = 188744404,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744405</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber27 = 188744405,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744406</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber28 = 188744406,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744407</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber29 = 188744407,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744408</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber30 = 188744408,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744409</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber31 = 188744409,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744410</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber32 = 188744410,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744411</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber33 = 188744411,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744412</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber34 = 188744412,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744413</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber35 = 188744413,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744414</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber36 = 188744414,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744415</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber37 = 188744415,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744416</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber38 = 188744416,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744417</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber39 = 188744417,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744418</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseNumber40 = 188744418,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744419</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode1 = 188744419,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744421</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode2 = 188744421,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744423</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode3 = 188744423,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744425</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode4 = 188744425,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744427</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode5 = 188744427,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744429</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode6 = 188744429,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744431</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode7 = 188744431,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744433</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode8 = 188744433,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744435</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode9 = 188744435,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744437</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode10 = 188744437,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744439</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode11 = 188744439,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744441</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode12 = 188744441,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744443</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode13 = 188744443,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744445</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode14 = 188744445,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744447</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode15 = 188744447,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744449</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode16 = 188744449,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744451</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode17 = 188744451,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744453</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode18 = 188744453,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744455</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode19 = 188744455,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744457</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode20 = 188744457,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744459</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode21 = 188744459,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744461</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode22 = 188744461,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744463</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode23 = 188744463,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744465</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode24 = 188744465,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744467</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode25 = 188744467,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744469</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode26 = 188744469,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744471</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode27 = 188744471,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744473</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode28 = 188744473,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744475</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode29 = 188744475,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744477</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseOutlineCode30 = 188744477,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744479</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText1 = 188744479,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744480</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText2 = 188744480,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744481</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText3 = 188744481,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744482</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText4 = 188744482,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744483</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText5 = 188744483,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744484</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText6 = 188744484,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744485</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText7 = 188744485,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744486</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText8 = 188744486,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744487</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText9 = 188744487,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744488</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText10 = 188744488,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744489</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText11 = 188744489,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744490</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText12 = 188744490,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744491</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText13 = 188744491,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744492</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText14 = 188744492,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744493</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText15 = 188744493,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744494</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText16 = 188744494,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744495</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText17 = 188744495,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744496</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText18 = 188744496,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744497</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText19 = 188744497,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744498</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText20 = 188744498,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744499</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText21 = 188744499,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744500</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText22 = 188744500,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744501</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText23 = 188744501,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744502</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText24 = 188744502,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744503</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText25 = 188744503,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744504</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText26 = 188744504,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744505</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText27 = 188744505,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744506</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText28 = 188744506,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744507</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText29 = 188744507,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744508</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText30 = 188744508,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744509</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText31 = 188744509,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744510</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText32 = 188744510,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744511</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText33 = 188744511,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744512</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText34 = 188744512,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744513</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText35 = 188744513,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744514</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText36 = 188744514,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744515</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText37 = 188744515,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744516</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText38 = 188744516,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744517</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText39 = 188744517,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744518</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseText40 = 188744518,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744519</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1DurationEstimated = 188744519,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744520</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2DurationEstimated = 188744520,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744521</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3DurationEstimated = 188744521,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744522</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4DurationEstimated = 188744522,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744523</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5DurationEstimated = 188744523,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744524</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6DurationEstimated = 188744524,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744525</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7DurationEstimated = 188744525,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744526</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8DurationEstimated = 188744526,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744527</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9DurationEstimated = 188744527,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744528</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10DurationEstimated = 188744528,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744529</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost1 = 188744529,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744530</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost2 = 188744530,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744531</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost3 = 188744531,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744532</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost4 = 188744532,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744533</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost5 = 188744533,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744534</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost6 = 188744534,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744535</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost7 = 188744535,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744536</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost8 = 188744536,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744537</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost9 = 188744537,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744538</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectCost10 = 188744538,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744539</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate1 = 188744539,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744540</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate2 = 188744540,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744541</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate3 = 188744541,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744542</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate4 = 188744542,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744543</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate5 = 188744543,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744544</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate6 = 188744544,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744545</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate7 = 188744545,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744546</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate8 = 188744546,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744547</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate9 = 188744547,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744548</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate10 = 188744548,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744549</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate11 = 188744549,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744550</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate12 = 188744550,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744551</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate13 = 188744551,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744552</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate14 = 188744552,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744553</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate15 = 188744553,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744554</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate16 = 188744554,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744555</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate17 = 188744555,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744556</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate18 = 188744556,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744557</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate19 = 188744557,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744558</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate20 = 188744558,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744559</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate21 = 188744559,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744560</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate22 = 188744560,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744561</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate23 = 188744561,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744562</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate24 = 188744562,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744563</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate25 = 188744563,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744564</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate26 = 188744564,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744565</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate27 = 188744565,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744566</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate28 = 188744566,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744567</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate29 = 188744567,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744568</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDate30 = 188744568,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744569</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration1 = 188744569,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744570</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration2 = 188744570,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744571</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration3 = 188744571,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744572</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration4 = 188744572,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744573</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration5 = 188744573,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744574</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration6 = 188744574,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744575</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration7 = 188744575,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744576</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration8 = 188744576,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744577</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration9 = 188744577,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744578</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectDuration10 = 188744578,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744589</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode1 = 188744589,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744590</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode2 = 188744590,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744591</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode3 = 188744591,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744592</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode4 = 188744592,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744593</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode5 = 188744593,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744594</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode6 = 188744594,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744595</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode7 = 188744595,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744596</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode8 = 188744596,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744597</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode9 = 188744597,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744598</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode10 = 188744598,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744599</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode11 = 188744599,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744600</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode12 = 188744600,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744601</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode13 = 188744601,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744602</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode14 = 188744602,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744603</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode15 = 188744603,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744604</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode16 = 188744604,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744605</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode17 = 188744605,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744606</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode18 = 188744606,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744607</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode19 = 188744607,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744608</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode20 = 188744608,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744609</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode21 = 188744609,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744610</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode22 = 188744610,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744611</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode23 = 188744611,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744612</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode24 = 188744612,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744613</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode25 = 188744613,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744614</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode26 = 188744614,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744615</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode27 = 188744615,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744616</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode28 = 188744616,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744617</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode29 = 188744617,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744618</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectOutlineCode30 = 188744618,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744649</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag1 = 188744649,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744650</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag2 = 188744650,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744651</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag3 = 188744651,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744652</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag4 = 188744652,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744653</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag5 = 188744653,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744654</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag6 = 188744654,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744655</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag7 = 188744655,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744656</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag8 = 188744656,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744657</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag9 = 188744657,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744658</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag10 = 188744658,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744659</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag11 = 188744659,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744660</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag12 = 188744660,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744661</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag13 = 188744661,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744662</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag14 = 188744662,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744663</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag15 = 188744663,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744664</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag16 = 188744664,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744665</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag17 = 188744665,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744666</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag18 = 188744666,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744667</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag19 = 188744667,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744668</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectFlag20 = 188744668,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744689</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber1 = 188744689,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744690</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber2 = 188744690,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744691</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber3 = 188744691,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744692</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber4 = 188744692,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744693</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber5 = 188744693,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744694</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber6 = 188744694,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744695</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber7 = 188744695,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744696</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber8 = 188744696,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744697</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber9 = 188744697,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744698</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber10 = 188744698,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744699</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber11 = 188744699,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744700</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber12 = 188744700,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744701</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber13 = 188744701,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744702</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber14 = 188744702,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744703</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber15 = 188744703,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744704</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber16 = 188744704,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744705</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber17 = 188744705,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744706</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber18 = 188744706,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744707</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber19 = 188744707,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744708</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber20 = 188744708,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744709</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber21 = 188744709,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744710</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber22 = 188744710,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744711</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber23 = 188744711,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744712</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber24 = 188744712,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744713</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber25 = 188744713,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744714</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber26 = 188744714,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744715</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber27 = 188744715,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744716</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber28 = 188744716,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744717</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber29 = 188744717,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744718</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber30 = 188744718,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744719</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber31 = 188744719,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744720</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber32 = 188744720,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744721</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber33 = 188744721,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744722</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber34 = 188744722,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744723</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber35 = 188744723,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744724</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber36 = 188744724,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744725</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber37 = 188744725,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744726</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber38 = 188744726,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744727</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber39 = 188744727,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744728</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectNumber40 = 188744728,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744729</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText1 = 188744729,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744730</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText2 = 188744730,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744731</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText3 = 188744731,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744732</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText4 = 188744732,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744733</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText5 = 188744733,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744734</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText6 = 188744734,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744735</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText7 = 188744735,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744736</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText8 = 188744736,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744737</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText9 = 188744737,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744738</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText10 = 188744738,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744739</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText11 = 188744739,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744740</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText12 = 188744740,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744741</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText13 = 188744741,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744742</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText14 = 188744742,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744743</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText15 = 188744743,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744744</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText16 = 188744744,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744745</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText17 = 188744745,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744746</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText18 = 188744746,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744747</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText19 = 188744747,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744748</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText20 = 188744748,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744749</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText21 = 188744749,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744750</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText22 = 188744750,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744751</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText23 = 188744751,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744752</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText24 = 188744752,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744753</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText25 = 188744753,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744754</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText26 = 188744754,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744755</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText27 = 188744755,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744756</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText28 = 188744756,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744757</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText29 = 188744757,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744758</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText30 = 188744758,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744759</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText31 = 188744759,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744760</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText32 = 188744760,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744761</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText33 = 188744761,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744762</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText34 = 188744762,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744763</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText35 = 188744763,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744764</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText36 = 188744764,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744765</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText37 = 188744765,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744766</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText38 = 188744766,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744767</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText39 = 188744767,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744768</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEnterpriseProjectText40 = 188744768,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744769</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode1 = 188744769,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744770</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode2 = 188744770,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744771</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode3 = 188744771,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744772</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode4 = 188744772,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744773</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode5 = 188744773,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744774</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode6 = 188744774,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744775</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode7 = 188744775,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744776</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode8 = 188744776,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744777</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode9 = 188744777,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744778</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode10 = 188744778,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744779</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode11 = 188744779,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744780</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode12 = 188744780,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744781</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode13 = 188744781,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744782</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode14 = 188744782,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744783</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode15 = 188744783,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744784</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode16 = 188744784,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744785</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode17 = 188744785,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744786</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode18 = 188744786,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744787</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode19 = 188744787,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744788</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode20 = 188744788,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744789</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode21 = 188744789,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744790</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode22 = 188744790,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744791</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode23 = 188744791,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744792</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode24 = 188744792,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744793</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode25 = 188744793,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744794</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode26 = 188744794,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744795</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode27 = 188744795,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744796</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode28 = 188744796,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744797</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseOutlineCode29 = 188744797,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744798</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseRBS = 188744798,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744799</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskPhysicalPercentComplete = 188744799,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744800</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDemandedRequested = 188744800,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744801</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStatusIndicator = 188744801,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744802</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskEarnedValueMethod = 188744802,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744809</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode20 = 188744809,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744810</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode21 = 188744810,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744811</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode22 = 188744811,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744812</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode23 = 188744812,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744813</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode24 = 188744813,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744814</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode25 = 188744814,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744815</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode26 = 188744815,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744816</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode27 = 188744816,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744817</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode28 = 188744817,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744818</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskResourceEnterpriseMultiValueCode29 = 188744818,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744819</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualWorkProtected = 188744819,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744820</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskActualOvertimeWorkProtected = 188744820,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520896</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceID = 205520896,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520897</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceName = 205520897,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520898</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceInitials = 205520898,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520899</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceGroup = 205520899,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520900</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceMaxUnits = 205520900,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520901</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseCalendar = 205520901,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520902</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStandardRate = 205520902,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520903</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOvertimeRate = 205520903,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520904</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText1 = 205520904,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520905</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText2 = 205520905,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520906</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCode = 205520906,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520907</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualCost = 205520907,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520908</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost = 205520908,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520909</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWork = 205520909,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520910</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualWork = 205520910,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520911</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineWork = 205520911,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520912</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOvertimeWork = 205520912,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520913</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineCost = 205520913,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520914</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCostPerUse = 205520914,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520915</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceAccrueAt = 205520915,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520916</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNotes = 205520916,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520917</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceRemainingCost = 205520917,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520918</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceRemainingWork = 205520918,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520919</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWorkVariance = 205520919,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520920</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCostVariance = 205520920,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520921</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOverallocated = 205520921,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520922</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourcePeakUnits = 205520922,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520923</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceUniqueID = 205520923,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520924</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceSheetNotes = 205520924,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520925</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourcePercentWorkComplete = 205520925,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520926</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText3 = 205520926,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520927</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText4 = 205520927,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520928</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText5 = 205520928,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520929</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceObjects = 205520929,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520930</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceLinkedFields = 205520930,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520931</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEMailAddress = 205520931,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520934</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceRegularWork = 205520934,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520935</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualOvertimeWork = 205520935,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520936</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceRemainingOvertimeWork = 205520936,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520943</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOvertimeCost = 205520943,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520944</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualOvertimeCost = 205520944,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520945</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceRemainingOvertimeCost = 205520945,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520947</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBCWS = 205520947,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520948</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBCWP = 205520948,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520949</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceACWP = 205520949,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520950</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceSV = 205520950,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520953</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceAvailableFrom = 205520953,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520954</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceAvailableTo = 205520954,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520982</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceIndicators = 205520982,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520993</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText6 = 205520993,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520994</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText7 = 205520994,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520995</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText8 = 205520995,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520996</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText9 = 205520996,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520997</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText10 = 205520997,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520998</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart1 = 205520998,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205520999</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart2 = 205520999,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521000</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart3 = 205521000,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521001</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart4 = 205521001,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521002</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart5 = 205521002,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521003</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish1 = 205521003,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521004</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish2 = 205521004,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521005</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish3 = 205521005,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521006</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish4 = 205521006,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521007</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish5 = 205521007,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521008</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber1 = 205521008,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521009</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber2 = 205521009,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521010</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber3 = 205521010,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521011</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber4 = 205521011,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521012</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber5 = 205521012,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521013</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration1 = 205521013,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521014</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration2 = 205521014,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521015</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration3 = 205521015,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521019</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost1 = 205521019,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521020</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost2 = 205521020,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521021</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost3 = 205521021,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521022</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag10 = 205521022,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521023</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag1 = 205521023,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521024</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag2 = 205521024,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521025</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag3 = 205521025,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521026</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag4 = 205521026,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521027</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag5 = 205521027,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521028</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag6 = 205521028,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521029</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag7 = 205521029,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521030</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag8 = 205521030,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521031</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag9 = 205521031,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521034</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceHyperlink = 205521034,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521035</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceHyperlinkAddress = 205521035,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521036</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceHyperlinkSubAddress = 205521036,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521037</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceHyperlinkHref = 205521037,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521040</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceIsAssignment = 205521040,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521055</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceTaskSummaryName = 205521055,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521059</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCanLevel = 205521059,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521060</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWorkContour = 205521060,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521062</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost4 = 205521062,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521063</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost5 = 205521063,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521064</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost6 = 205521064,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521065</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost7 = 205521065,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521066</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost8 = 205521066,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521067</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost9 = 205521067,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521068</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCost10 = 205521068,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521069</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate1 = 205521069,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521070</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate2 = 205521070,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521071</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate3 = 205521071,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521072</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate4 = 205521072,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521073</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate5 = 205521073,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521074</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate6 = 205521074,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521075</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate7 = 205521075,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521076</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate8 = 205521076,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521077</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate9 = 205521077,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521078</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDate10 = 205521078,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521079</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration4 = 205521079,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521080</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration5 = 205521080,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521081</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration6 = 205521081,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521082</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration7 = 205521082,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521083</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration8 = 205521083,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521084</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration9 = 205521084,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521085</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDuration10 = 205521085,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521086</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish6 = 205521086,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521087</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish7 = 205521087,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521088</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish8 = 205521088,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521089</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish9 = 205521089,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521090</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish10 = 205521090,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521091</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag11 = 205521091,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521092</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag12 = 205521092,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521093</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag13 = 205521093,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521094</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag14 = 205521094,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521095</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag15 = 205521095,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521096</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag16 = 205521096,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521097</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag17 = 205521097,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521098</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag18 = 205521098,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521099</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag19 = 205521099,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521100</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFlag20 = 205521100,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521101</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber6 = 205521101,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521102</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber7 = 205521102,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521103</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber8 = 205521103,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521104</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber9 = 205521104,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521105</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber10 = 205521105,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521106</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber11 = 205521106,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521107</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber12 = 205521107,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521108</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber13 = 205521108,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521109</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber14 = 205521109,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521110</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber15 = 205521110,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521111</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber16 = 205521111,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521112</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber17 = 205521112,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521113</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber18 = 205521113,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521114</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber19 = 205521114,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521115</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceNumber20 = 205521115,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521116</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart6 = 205521116,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521117</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart7 = 205521117,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521118</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart8 = 205521118,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521119</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart9 = 205521119,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521120</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart10 = 205521120,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521121</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText11 = 205521121,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521122</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText12 = 205521122,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521123</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText13 = 205521123,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521124</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText14 = 205521124,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521125</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText15 = 205521125,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521126</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText16 = 205521126,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521127</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText17 = 205521127,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521128</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText18 = 205521128,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521129</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText19 = 205521129,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521130</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText20 = 205521130,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521131</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText21 = 205521131,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521132</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText22 = 205521132,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521133</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText23 = 205521133,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521134</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText24 = 205521134,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521135</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText25 = 205521135,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521136</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText26 = 205521136,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521137</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText27 = 205521137,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521138</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText28 = 205521138,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521139</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText29 = 205521139,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521140</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceText30 = 205521140,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521148</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourcePhonetics = 205521148,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521149</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceIndex = 205521149,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521153</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceAssignmentDelay = 205521153,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521154</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceAssignmentUnits = 205521154,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521155</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineStart = 205521155,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521156</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineFinish = 205521156,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521157</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceConfirmed = 205521157,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521158</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceFinish = 205521158,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521159</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceLevelingDelay = 205521159,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521160</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceResponsePending = 205521160,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521161</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceStart = 205521161,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521162</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceTeamStatusPending = 205521162,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521163</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceUpdateNeeded = 205521163,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521164</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCV = 205521164,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521165</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCostRateTable = 205521165,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521168</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWorkgroup = 205521168,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521169</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceProject = 205521169,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521174</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode1 = 205521174,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521176</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode2 = 205521176,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521178</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode3 = 205521178,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521180</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode4 = 205521180,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521182</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode5 = 205521182,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521184</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode6 = 205521184,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521186</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode7 = 205521186,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521188</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode8 = 205521188,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521190</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode9 = 205521190,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521192</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceOutlineCode10 = 205521192,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521195</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceMaterialLabel = 205521195,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521196</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceType = 205521196,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521197</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceVAC = 205521197,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521202</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceGroupbySummary = 205521202,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521207</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWindowsUserAccount = 205521207,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521208</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceHyperlinkScreenTip = 205521208,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521238</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1Work = 205521238,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521239</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1Cost = 205521239,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521244</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1Start = 205521244,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521245</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1Finish = 205521245,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521248</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2Work = 205521248,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521249</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2Cost = 205521249,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521254</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2Start = 205521254,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521255</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2Finish = 205521255,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521258</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3Work = 205521258,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521259</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3Cost = 205521259,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521264</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3Start = 205521264,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521265</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3Finish = 205521265,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521268</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4Work = 205521268,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521269</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4Cost = 205521269,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521274</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4Start = 205521274,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521275</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4Finish = 205521275,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521278</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5Work = 205521278,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521279</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5Cost = 205521279,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521284</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5Start = 205521284,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521285</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5Finish = 205521285,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521288</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6Work = 205521288,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521289</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6Cost = 205521289,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521294</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6Start = 205521294,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521295</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6Finish = 205521295,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521298</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7Work = 205521298,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521299</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7Cost = 205521299,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521304</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7Start = 205521304,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521305</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7Finish = 205521305,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521308</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8Work = 205521308,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521309</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8Cost = 205521309,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521314</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8Start = 205521314,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521315</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8Finish = 205521315,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521318</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9Work = 205521318,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521319</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9Cost = 205521319,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521324</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9Start = 205521324,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521325</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9Finish = 205521325,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521328</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10Work = 205521328,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521329</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10Cost = 205521329,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521334</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10Start = 205521334,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521335</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10Finish = 205521335,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521339</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseUniqueID = 205521339,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521342</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost1 = 205521342,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521343</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost2 = 205521343,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521344</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost3 = 205521344,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521345</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost4 = 205521345,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521346</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost5 = 205521346,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521347</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost6 = 205521347,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521348</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost7 = 205521348,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521349</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost8 = 205521349,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521350</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost9 = 205521350,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521351</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCost10 = 205521351,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521352</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate1 = 205521352,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521353</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate2 = 205521353,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521354</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate3 = 205521354,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521355</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate4 = 205521355,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521356</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate5 = 205521356,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521357</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate6 = 205521357,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521358</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate7 = 205521358,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521359</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate8 = 205521359,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521360</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate9 = 205521360,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521361</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate10 = 205521361,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521362</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate11 = 205521362,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521363</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate12 = 205521363,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521364</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate13 = 205521364,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521365</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate14 = 205521365,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521366</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate15 = 205521366,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521367</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate16 = 205521367,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521368</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate17 = 205521368,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521369</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate18 = 205521369,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521370</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate19 = 205521370,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521371</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate20 = 205521371,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521372</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate21 = 205521372,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521373</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate22 = 205521373,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521374</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate23 = 205521374,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521375</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate24 = 205521375,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521376</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate25 = 205521376,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521377</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate26 = 205521377,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521378</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate27 = 205521378,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521379</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate28 = 205521379,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521380</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate29 = 205521380,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521381</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDate30 = 205521381,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521382</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration1 = 205521382,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521383</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration2 = 205521383,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521384</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration3 = 205521384,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521385</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration4 = 205521385,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521386</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration5 = 205521386,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521387</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration6 = 205521387,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521388</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration7 = 205521388,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521389</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration8 = 205521389,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521390</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration9 = 205521390,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521391</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseDuration10 = 205521391,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521402</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag1 = 205521402,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521403</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag2 = 205521403,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521404</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag3 = 205521404,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521405</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag4 = 205521405,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521406</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag5 = 205521406,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521407</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag6 = 205521407,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521408</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag7 = 205521408,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521409</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag8 = 205521409,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521410</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag9 = 205521410,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521411</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag10 = 205521411,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521412</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag11 = 205521412,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521413</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag12 = 205521413,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521414</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag13 = 205521414,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521415</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag14 = 205521415,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521416</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag15 = 205521416,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521417</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag16 = 205521417,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521418</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag17 = 205521418,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521419</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag18 = 205521419,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521420</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag19 = 205521420,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521421</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseFlag20 = 205521421,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521442</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber1 = 205521442,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521443</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber2 = 205521443,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521444</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber3 = 205521444,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521445</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber4 = 205521445,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521446</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber5 = 205521446,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521447</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber6 = 205521447,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521448</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber7 = 205521448,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521449</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber8 = 205521449,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521450</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber9 = 205521450,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521451</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber10 = 205521451,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521452</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber11 = 205521452,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521453</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber12 = 205521453,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521454</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber13 = 205521454,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521455</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber14 = 205521455,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521456</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber15 = 205521456,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521457</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber16 = 205521457,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521458</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber17 = 205521458,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521459</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber18 = 205521459,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521460</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber19 = 205521460,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521461</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber20 = 205521461,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521462</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber21 = 205521462,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521463</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber22 = 205521463,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521464</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber23 = 205521464,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521465</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber24 = 205521465,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521466</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber25 = 205521466,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521467</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber26 = 205521467,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521468</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber27 = 205521468,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521469</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber28 = 205521469,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521470</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber29 = 205521470,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521471</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber30 = 205521471,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521472</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber31 = 205521472,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521473</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber32 = 205521473,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521474</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber33 = 205521474,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521475</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber34 = 205521475,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521476</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber35 = 205521476,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521477</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber36 = 205521477,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521478</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber37 = 205521478,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521479</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber38 = 205521479,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521480</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber39 = 205521480,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521481</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNumber40 = 205521481,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521482</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode1 = 205521482,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521484</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode2 = 205521484,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521486</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode3 = 205521486,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521488</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode4 = 205521488,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521490</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode5 = 205521490,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521492</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode6 = 205521492,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521494</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode7 = 205521494,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521496</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode8 = 205521496,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521498</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode9 = 205521498,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521500</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode10 = 205521500,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521502</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode11 = 205521502,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521504</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode12 = 205521504,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521506</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode13 = 205521506,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521508</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode14 = 205521508,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521510</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode15 = 205521510,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521512</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode16 = 205521512,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521514</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode17 = 205521514,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521516</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode18 = 205521516,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521518</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode19 = 205521518,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521520</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode20 = 205521520,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521522</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode21 = 205521522,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521524</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode22 = 205521524,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521526</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode23 = 205521526,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521528</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode24 = 205521528,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521530</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode25 = 205521530,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521532</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode26 = 205521532,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521534</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode27 = 205521534,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521536</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode28 = 205521536,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521538</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseOutlineCode29 = 205521538,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521540</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseRBS = 205521540,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521542</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText1 = 205521542,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521543</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText2 = 205521543,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521544</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText3 = 205521544,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521545</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText4 = 205521545,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521546</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText5 = 205521546,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521547</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText6 = 205521547,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521548</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText7 = 205521548,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521549</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText8 = 205521549,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521550</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText9 = 205521550,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521551</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText10 = 205521551,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521552</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText11 = 205521552,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521553</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText12 = 205521553,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521554</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText13 = 205521554,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521555</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText14 = 205521555,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521556</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText15 = 205521556,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521557</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText16 = 205521557,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521558</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText17 = 205521558,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521559</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText18 = 205521559,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521560</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText19 = 205521560,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521561</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText20 = 205521561,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521562</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText21 = 205521562,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521563</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText22 = 205521563,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521564</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText23 = 205521564,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521565</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText24 = 205521565,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521566</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText25 = 205521566,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521567</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText26 = 205521567,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521568</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText27 = 205521568,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521569</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText28 = 205521569,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521570</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText29 = 205521570,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521571</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText30 = 205521571,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521572</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText31 = 205521572,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521573</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText32 = 205521573,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521574</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText33 = 205521574,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521575</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText34 = 205521575,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521576</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText35 = 205521576,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521577</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText36 = 205521577,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521578</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText37 = 205521578,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521579</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText38 = 205521579,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521580</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText39 = 205521580,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521581</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseText40 = 205521581,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521582</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseGeneric = 205521582,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521583</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseBaseCalendar = 205521583,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521584</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseRequiredValues = 205521584,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521585</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseNameUsed = 205521585,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521586</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDemandedRequested = 205521586,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521587</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterprise = 205521587,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521588</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseIsCheckedOut = 205521588,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521589</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseCheckedOutBy = 205521589,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521590</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseLastModifiedDate = 205521590,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521591</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseTeamMember = 205521591,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521592</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseInactive = 205521592,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521595</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBookingType = 205521595,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521596</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue20 = 205521596,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521598</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue21 = 205521598,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521600</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue22 = 205521600,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521602</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue23 = 205521602,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521604</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue24 = 205521604,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521606</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue25 = 205521606,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521608</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue26 = 205521608,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521610</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue27 = 205521610,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521612</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue28 = 205521612,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521614</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceEnterpriseMultiValue29 = 205521614,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521616</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualWorkProtected = 205521616,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521617</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceActualOvertimeWorkProtected = 205521617,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521622</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCreated = 205521622,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188743700</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDelay = 188743700,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744823</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskGuid = 188744823,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744824</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskCalendarGuid = 188744824,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744826</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeliverableGuid = 188744826,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744827</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeliverableType = 188744827,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744832</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeliverableStart = 188744832,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744833</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeliverableFinish = 188744833,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744845</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskIsPublished = 188744845,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744846</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskStatusManagerName = 188744846,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744847</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskErrorMessage = 188744847,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744851</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBudgetWork = 188744851,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744852</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBudgetCost = 188744852,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744853</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineFixedCostAccrual = 188744853,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744854</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineDeliverableStart = 188744854,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744855</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineDeliverableFinish = 188744855,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744856</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineBudgetWork = 188744856,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744857</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaselineBudgetCost = 188744857,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744860</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1FixedCostAccrual = 188744860,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744861</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1DeliverableStart = 188744861,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744862</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1DeliverableFinish = 188744862,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744863</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1BudgetWork = 188744863,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744864</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline1BudgetCost = 188744864,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744867</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2FixedCostAccrual = 188744867,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744868</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2DeliverableStart = 188744868,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744869</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2DeliverableFinish = 188744869,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744870</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2BudgetWork = 188744870,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744871</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline2BudgetCost = 188744871,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744874</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3FixedCostAccrual = 188744874,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744875</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3DeliverableStart = 188744875,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744876</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3DeliverableFinish = 188744876,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744877</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3BudgetWork = 188744877,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744878</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline3BudgetCost = 188744878,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744881</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4FixedCostAccrual = 188744881,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744882</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4DeliverableStart = 188744882,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744883</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4DeliverableFinish = 188744883,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744884</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4BudgetWork = 188744884,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744885</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline4BudgetCost = 188744885,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744888</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5FixedCostAccrual = 188744888,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744889</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5DeliverableStart = 188744889,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744890</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5DeliverableFinish = 188744890,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744891</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5BudgetWork = 188744891,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744892</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline5BudgetCost = 188744892,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744895</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6FixedCostAccrual = 188744895,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744896</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6DeliverableStart = 188744896,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744897</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6DeliverableFinish = 188744897,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744898</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6BudgetWork = 188744898,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744899</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline6BudgetCost = 188744899,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744902</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7FixedCostAccrual = 188744902,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744903</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7DeliverableStart = 188744903,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744904</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7DeliverableFinish = 188744904,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744905</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7BudgetWork = 188744905,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744906</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline7BudgetCost = 188744906,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744909</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8FixedCostAccrual = 188744909,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744910</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8DeliverableStart = 188744910,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744911</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8DeliverableFinish = 188744911,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744912</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8BudgetWork = 188744912,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744913</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline8BudgetCost = 188744913,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744916</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9FixedCostAccrual = 188744916,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744917</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9DeliverableStart = 188744917,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744918</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9DeliverableFinish = 188744918,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744919</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9BudgetWork = 188744919,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744920</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline9BudgetCost = 188744920,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744923</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10FixedCostAccrual = 188744923,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744924</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10DeliverableStart = 188744924,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744925</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10DeliverableFinish = 188744925,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744926</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10BudgetWork = 188744926,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744927</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskBaseline10BudgetCost = 188744927,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744930</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskRecalcFlags = 188744930,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>188744956</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjTaskDeliverableName = 188744956,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521236</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceWBS = 205521236,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521624</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceGuid = 205521624,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521625</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCalendarGuid = 205521625,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521634</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceErrorMessage = 205521634,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521636</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceDefaultAssignmentOwner = 205521636,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521648</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBudget = 205521648,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521649</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBudgetWork = 205521649,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521650</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBudgetCost = 205521650,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521651</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjImportResource = 205521651,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521652</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineBudgetWork = 205521652,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521653</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaselineBudgetCost = 205521653,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521656</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1BudgetWork = 205521656,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521657</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline1BudgetCost = 205521657,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521660</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2BudgetWork = 205521660,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521661</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline2BudgetCost = 205521661,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521664</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3BudgetWork = 205521664,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521665</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline3BudgetCost = 205521665,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521668</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4BudgetWork = 205521668,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521669</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline4BudgetCost = 205521669,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521672</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5BudgetWork = 205521672,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521673</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline5BudgetCost = 205521673,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521676</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6BudgetWork = 205521676,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521677</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline6BudgetCost = 205521677,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521680</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7BudgetWork = 205521680,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521681</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline7BudgetCost = 205521681,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521684</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8BudgetWork = 205521684,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521685</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline8BudgetCost = 205521685,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521688</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9BudgetWork = 205521688,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521689</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline9BudgetCost = 205521689,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521692</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10BudgetWork = 205521692,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521693</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceBaseline10BudgetCost = 205521693,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521696</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceIsTeam = 205521696,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 16
		 /// </summary>
		 /// <remarks>205521697</remarks>
		 [SupportByVersion("MSProject", 11,12,14,16)]
		 pjResourceCostCenter = 205521697,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744959</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskActive = 188744959,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744960</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskManual = 188744960,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744961</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskPlaceholder = 188744961,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744962</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskWarning = 188744962,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744965</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskStartText = 188744965,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744966</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskFinishText = 188744966,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744967</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskDurationText = 188744967,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744975</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskIsStartValid = 188744975,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744976</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskIsFinishValid = 188744976,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744977</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskIsDurationValid = 188744977,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744979</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaselineStartText = 188744979,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744980</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaselineFinishText = 188744980,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744981</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaselineDurationText = 188744981,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744982</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline1StartText = 188744982,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744983</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline1FinishText = 188744983,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744984</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline1DurationText = 188744984,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744985</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline2StartText = 188744985,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744986</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline2FinishText = 188744986,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744987</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline2DurationText = 188744987,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744988</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline3StartText = 188744988,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744989</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline3FinishText = 188744989,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744990</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline3DurationText = 188744990,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744991</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline4StartText = 188744991,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744992</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline4FinishText = 188744992,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744993</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline4DurationText = 188744993,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744994</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline5StartText = 188744994,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744995</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline5FinishText = 188744995,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744996</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline5DurationText = 188744996,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744997</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline6StartText = 188744997,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744998</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline6FinishText = 188744998,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188744999</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline6DurationText = 188744999,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745000</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline7StartText = 188745000,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745001</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline7FinishText = 188745001,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745002</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline7DurationText = 188745002,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745003</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline8StartText = 188745003,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745004</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline8FinishText = 188745004,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745005</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline8DurationText = 188745005,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745006</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline9StartText = 188745006,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745007</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline9FinishText = 188745007,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745008</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline9DurationText = 188745008,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745009</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline10StartText = 188745009,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745010</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline10FinishText = 188745010,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745011</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskBaseline10DurationText = 188745011,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745012</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskIgnoreWarnings = 188745012,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745015</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskAssignmentPeakUnits = 188745015,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745018</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskScheduledStart = 188745018,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745019</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskScheduledFinish = 188745019,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>188745020</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjTaskScheduledDuration = 188745020,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>188745061</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjTaskPathDrivingPredecessor = 188745061,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>188745062</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjTaskPathPredecessor = 188745062,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>188745063</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjTaskPathDrivenSuccessor = 188745063,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>188745064</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjTaskPathSuccessor = 188745064,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>188745084</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjTaskCompliant = 188745084,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521737</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceIsLocked = 205521737,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521738</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceCompliant = 205521738,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521755</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementDraftMaxUnits = 205521755,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521756</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementProposedMaxUnits = 205521756,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521757</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementCommittedMaxUnits = 205521757,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521759</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementDraftWork = 205521759,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521761</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementProposedWork = 205521761,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521763</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementCommittedWork = 205521763,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521765</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementDraftStart = 205521765,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521766</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementProposedStart = 205521766,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521767</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementCommittedStart = 205521767,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521768</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementDraftFinish = 205521768,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521769</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementProposedFinish = 205521769,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>205521770</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjResourceEngagementCommittedFinish = 205521770,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401921</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementGuid = 440401921,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401922</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementName = 440401922,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401923</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProjectGuid = 440401923,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401924</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProjectName = 440401924,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401927</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementResourceGuid = 440401927,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401928</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementResourceName = 440401928,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401929</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementResourceID = 440401929,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401930</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementModifiedByGuid = 440401930,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401931</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementModifiedByName = 440401931,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401932</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementSubmittedByGuid = 440401932,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401933</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementSubmittedByName = 440401933,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401934</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementReviewedByGuid = 440401934,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401935</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementReviewedByName = 440401935,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401936</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementCreatedDate = 440401936,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401937</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementSubmittedDate = 440401937,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401938</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementReviewedDate = 440401938,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401939</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementModifiedDate = 440401939,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401946</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementDraftMaxUnits = 440401946,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401947</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProposedMaxUnits = 440401947,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401948</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementCommittedMaxUnits = 440401948,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401949</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementDraftWork = 440401949,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401951</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProposedWork = 440401951,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401953</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementCommittedWork = 440401953,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401955</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementStatus = 440401955,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401968</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementIndex = 440401968,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401993</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementDraftStart = 440401993,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401994</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProposedStart = 440401994,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401995</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementCommittedStart = 440401995,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401996</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementDraftFinish = 440401996,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401997</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementProposedFinish = 440401997,

		 /// <summary>
		 /// SupportByVersion MSProject 16
		 /// </summary>
		 /// <remarks>440401998</remarks>
		 [SupportByVersion("MSProject", 16)]
		 pjEngagementCommittedFinish = 440401998
	}
}