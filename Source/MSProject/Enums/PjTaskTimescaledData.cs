using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjTaskTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledWork = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledRegularWork = 166,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledOvertimeWork = 163,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualWork = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualOvertimeWork = 164,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativeWork = 176,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineWork = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledOverallocation = 167,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCost = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledFixedCost = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualCost = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaselineCost = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativeCost = 177,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBCWS = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBCWP = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledACWP = 120,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledSV = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCV = 83,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledPercentComplete = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCumulativePercentComplete = 240,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>485</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1Work = 485,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>484</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline1Cost = 484,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>496</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2Work = 496,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>495</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline2Cost = 495,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>507</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3Work = 507,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline3Cost = 506,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>518</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4Work = 518,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>517</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline4Cost = 517,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>529</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5Work = 529,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>528</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline5Cost = 528,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>547</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6Work = 547,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>546</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline6Cost = 546,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>558</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7Work = 558,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>557</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline7Cost = 557,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>569</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8Work = 569,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>568</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline8Cost = 568,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>580</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9Work = 580,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>579</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline9Cost = 579,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>591</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10Work = 591,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>590</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledBaseline10Cost = 590,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>537</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCPI = 537,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>538</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledSPI = 538,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>539</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledCVP = 539,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>540</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledSVP = 540,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualFixedCost = 171,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1139</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualWorkProtected = 1139,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1140</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjTaskTimescaledActualOvertimeWorkProtected = 1140,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1172</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBudgetWork = 1172,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1173</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBudgetCost = 1173,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1177</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaselineBudgetWork = 1177,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1178</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaselineBudgetCost = 1178,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1184</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline1BudgetWork = 1184,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1185</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline1BudgetCost = 1185,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1191</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline2BudgetWork = 1191,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1192</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline2BudgetCost = 1192,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1198</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline3BudgetWork = 1198,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1199</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline3BudgetCost = 1199,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1205</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline4BudgetWork = 1205,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1206</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline4BudgetCost = 1206,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1212</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline5BudgetWork = 1212,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1213</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline5BudgetCost = 1213,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1219</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline6BudgetWork = 1219,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1220</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline6BudgetCost = 1220,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1226</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline7BudgetWork = 1226,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1227</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline7BudgetCost = 1227,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1233</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline8BudgetWork = 1233,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1234</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline8BudgetCost = 1234,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1240</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline9BudgetWork = 1240,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1241</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline9BudgetCost = 1241,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1247</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline10BudgetWork = 1247,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1248</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjTaskTimescaledBaseline10BudgetCost = 1248
	}
}