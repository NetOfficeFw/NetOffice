using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867715(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjTaskTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledWork = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledRegularWork = 166,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledOvertimeWork = 163,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualWork = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualOvertimeWork = 164,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativeWork = 176,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineWork = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledOverallocation = 167,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCost = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledFixedCost = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualCost = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineCost = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativeCost = 177,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBCWS = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBCWP = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledACWP = 120,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledSV = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCV = 83,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledPercentComplete = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativePercentComplete = 240,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>485</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1Work = 485,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>484</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1Cost = 484,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>496</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2Work = 496,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>495</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2Cost = 495,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>507</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3Work = 507,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3Cost = 506,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>518</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4Work = 518,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>517</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4Cost = 517,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>529</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5Work = 529,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>528</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5Cost = 528,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>547</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6Work = 547,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>546</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6Cost = 546,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>558</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7Work = 558,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>557</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7Cost = 557,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>569</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8Work = 569,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>568</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8Cost = 568,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>580</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9Work = 580,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>579</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9Cost = 579,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>591</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10Work = 591,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>590</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10Cost = 590,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>537</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCPI = 537,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>538</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledSPI = 538,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>539</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledCVP = 539,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>540</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledSVP = 540,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualFixedCost = 171,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1139</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualWorkProtected = 1139,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1140</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledActualOvertimeWorkProtected = 1140,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1172</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBudgetWork = 1172,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1173</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBudgetCost = 1173,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1177</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineBudgetWork = 1177,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1178</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineBudgetCost = 1178,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1184</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1BudgetWork = 1184,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1185</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1BudgetCost = 1185,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1191</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2BudgetWork = 1191,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1192</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2BudgetCost = 1192,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1198</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3BudgetWork = 1198,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1199</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3BudgetCost = 1199,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1205</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4BudgetWork = 1205,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1206</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4BudgetCost = 1206,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1212</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5BudgetWork = 1212,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1213</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5BudgetCost = 1213,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1219</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6BudgetWork = 1219,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1220</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6BudgetCost = 1220,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1226</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7BudgetWork = 1226,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1227</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7BudgetCost = 1227,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1233</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8BudgetWork = 1233,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1234</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8BudgetCost = 1234,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1240</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9BudgetWork = 1240,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1241</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9BudgetCost = 1241,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1247</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10BudgetWork = 1247,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1248</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10BudgetCost = 1248,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1341</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledCumulativeActualWork = 1341,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1342</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledRemainingCumulativeActualWork = 1342,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1343</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledRemainingCumulativeWork = 1343,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1344</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledRemainingTasks = 1344,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1345</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledRemainingActualTasks = 1345,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1346</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaselineRemainingCumulativeWork = 1346,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1347</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline1RemainingCumulativeWork = 1347,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1348</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline2RemainingCumulativeWork = 1348,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1349</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline3RemainingCumulativeWork = 1349,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1350</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline4RemainingCumulativeWork = 1350,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1351</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline5RemainingCumulativeWork = 1351,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1352</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline6RemainingCumulativeWork = 1352,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1353</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline7RemainingCumulativeWork = 1353,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1354</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline8RemainingCumulativeWork = 1354,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1355</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline9RemainingCumulativeWork = 1355,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1356</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline10RemainingCumulativeWork = 1356,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1357</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaselineRemainingTasks = 1357,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1358</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline1RemainingTasks = 1358,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1359</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline2RemainingTasks = 1359,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1360</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline3RemainingTasks = 1360,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline4RemainingTasks = 1361,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1362</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline5RemainingTasks = 1362,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1363</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline6RemainingTasks = 1363,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1364</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline7RemainingTasks = 1364,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1365</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline8RemainingTasks = 1365,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1366</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline9RemainingTasks = 1366,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1367</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline10RemainingTasks = 1367,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1368</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaselineCumulativeWork = 1368,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1369</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline1CumulativeWork = 1369,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1370</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline2CumulativeWork = 1370,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1371</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline3CumulativeWork = 1371,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1372</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline4CumulativeWork = 1372,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1373</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline5CumulativeWork = 1373,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1374</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline6CumulativeWork = 1374,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1375</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline7CumulativeWork = 1375,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1376</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline8CumulativeWork = 1376,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1377</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline9CumulativeWork = 1377,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1378</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjTaskTimescaledBaseline10CumulativeWork = 1378
	}
}