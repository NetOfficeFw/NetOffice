using System;
using LateBindingApi.Core;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjResourceTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledWork = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledRegularWork = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledOvertimeWork = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledActualWork = 14,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledActualOvertimeWork = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledCumulativeWork = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaselineWork = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledOverallocation = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledPercentAllocation = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledPeakUnits = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledRemainingAvailability = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledCost = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledActualCost = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaselineCost = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledCumulativeCost = 50,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBCWS = 51,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBCWP = 52,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledACWP = 53,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledSV = 54,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>268</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledCV = 268,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledWorkAvailability = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledUnitAvailability = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>342</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline1Work = 342,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>343</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline1Cost = 343,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>352</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline2Work = 352,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline2Cost = 353,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>362</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline3Work = 362,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>363</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline3Cost = 363,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>372</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline4Work = 372,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>373</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline4Cost = 373,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>382</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline5Work = 382,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>383</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline5Cost = 383,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>392</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline6Work = 392,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>393</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline6Cost = 393,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>402</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline7Work = 402,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>403</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline7Cost = 403,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>412</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline8Work = 412,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>413</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline8Cost = 413,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>422</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline9Work = 422,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>423</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline9Cost = 423,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>432</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline10Work = 432,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>433</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline10Cost = 433,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>720</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledActualWorkProtected = 720,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>721</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledActualOvertimeWorkProtected = 721,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>754</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBudgetWork = 754,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>755</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBudgetCost = 755,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>757</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaselineBudgetWork = 757,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>758</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaselineBudgetCost = 758,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>761</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline1BudgetWork = 761,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>762</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline1BudgetCost = 762,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>765</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline2BudgetWork = 765,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>766</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline2BudgetCost = 766,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>769</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline3BudgetWork = 769,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>770</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline3BudgetCost = 770,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>773</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline4BudgetWork = 773,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>774</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline4BudgetCost = 774,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>777</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline5BudgetWork = 777,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>778</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline5BudgetCost = 778,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>781</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline6BudgetWork = 781,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>782</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline6BudgetCost = 782,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>785</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline7BudgetWork = 785,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>786</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline7BudgetCost = 786,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>789</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline8BudgetWork = 789,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>790</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline8BudgetCost = 790,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>793</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline9BudgetWork = 793,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>794</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline9BudgetCost = 794,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>797</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline10BudgetWork = 797,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>798</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjResourceTimescaledBaseline10BudgetCost = 798
	}
}