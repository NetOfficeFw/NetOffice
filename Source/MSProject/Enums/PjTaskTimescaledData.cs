using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjTaskTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledWork = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledRegularWork = 166,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledOvertimeWork = 163,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualWork = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualOvertimeWork = 164,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCumulativeWork = 176,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaselineWork = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledOverallocation = 167,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCost = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledFixedCost = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualCost = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaselineCost = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCumulativeCost = 177,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBCWS = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBCWP = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledACWP = 120,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledSV = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCV = 83,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledPercentComplete = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCumulativePercentComplete = 240,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>485</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline1Work = 485,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>484</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline1Cost = 484,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>496</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline2Work = 496,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>495</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline2Cost = 495,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>507</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline3Work = 507,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline3Cost = 506,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>518</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline4Work = 518,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>517</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline4Cost = 517,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>529</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline5Work = 529,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>528</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline5Cost = 528,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>547</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline6Work = 547,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>546</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline6Cost = 546,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>558</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline7Work = 558,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>557</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline7Cost = 557,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>569</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline8Work = 569,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>568</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline8Cost = 568,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>580</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline9Work = 580,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>579</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline9Cost = 579,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>591</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline10Work = 591,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>590</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledBaseline10Cost = 590,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>537</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCPI = 537,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>538</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledSPI = 538,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>539</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledCVP = 539,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>540</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledSVP = 540,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualFixedCost = 171,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1139</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualWorkProtected = 1139,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1140</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTaskTimescaledActualOvertimeWorkProtected = 1140,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1172</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBudgetWork = 1172,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1173</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBudgetCost = 1173,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1177</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaselineBudgetWork = 1177,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1178</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaselineBudgetCost = 1178,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1184</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline1BudgetWork = 1184,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1185</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline1BudgetCost = 1185,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1191</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline2BudgetWork = 1191,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1192</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline2BudgetCost = 1192,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1198</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline3BudgetWork = 1198,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1199</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline3BudgetCost = 1199,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1205</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline4BudgetWork = 1205,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1206</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline4BudgetCost = 1206,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1212</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline5BudgetWork = 1212,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1213</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline5BudgetCost = 1213,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1219</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline6BudgetWork = 1219,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1220</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline6BudgetCost = 1220,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1226</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline7BudgetWork = 1226,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1227</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline7BudgetCost = 1227,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1233</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline8BudgetWork = 1233,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1234</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline8BudgetCost = 1234,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1240</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline9BudgetWork = 1240,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1241</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline9BudgetCost = 1241,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1247</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline10BudgetWork = 1247,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1248</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjTaskTimescaledBaseline10BudgetCost = 1248,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1341</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledCumulativeActualWork = 1341,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1342</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledRemainingCumulativeActualWork = 1342,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1343</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledRemainingCumulativeWork = 1343,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1344</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledRemainingTasks = 1344,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1345</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledRemainingActualTasks = 1345,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1346</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaselineRemainingCumulativeWork = 1346,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1347</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline1RemainingCumulativeWork = 1347,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1348</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline2RemainingCumulativeWork = 1348,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1349</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline3RemainingCumulativeWork = 1349,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1350</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline4RemainingCumulativeWork = 1350,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1351</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline5RemainingCumulativeWork = 1351,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1352</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline6RemainingCumulativeWork = 1352,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1353</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline7RemainingCumulativeWork = 1353,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1354</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline8RemainingCumulativeWork = 1354,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1355</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline9RemainingCumulativeWork = 1355,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1356</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline10RemainingCumulativeWork = 1356,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1357</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaselineRemainingTasks = 1357,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1358</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline1RemainingTasks = 1358,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1359</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline2RemainingTasks = 1359,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1360</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline3RemainingTasks = 1360,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline4RemainingTasks = 1361,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1362</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline5RemainingTasks = 1362,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1363</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline6RemainingTasks = 1363,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1364</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline7RemainingTasks = 1364,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1365</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline8RemainingTasks = 1365,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1366</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline9RemainingTasks = 1366,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1367</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline10RemainingTasks = 1367,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1368</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaselineCumulativeWork = 1368,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1369</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline1CumulativeWork = 1369,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1370</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline2CumulativeWork = 1370,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1371</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline3CumulativeWork = 1371,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1372</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline4CumulativeWork = 1372,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1373</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline5CumulativeWork = 1373,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1374</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline6CumulativeWork = 1374,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1375</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline7CumulativeWork = 1375,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1376</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline8CumulativeWork = 1376,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1377</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline9CumulativeWork = 1377,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1378</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjTaskTimescaledBaseline10CumulativeWork = 1378
	}
}