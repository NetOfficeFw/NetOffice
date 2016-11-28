using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866545(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjResourceTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledWork = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledRegularWork = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledOvertimeWork = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledActualWork = 14,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledActualOvertimeWork = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledCumulativeWork = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaselineWork = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledOverallocation = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledPercentAllocation = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledPeakUnits = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledRemainingAvailability = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledCost = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledActualCost = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaselineCost = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledCumulativeCost = 50,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBCWS = 51,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBCWP = 52,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledACWP = 53,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledSV = 54,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>268</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledCV = 268,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledWorkAvailability = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledUnitAvailability = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>342</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline1Work = 342,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>343</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline1Cost = 343,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>352</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline2Work = 352,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline2Cost = 353,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>362</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline3Work = 362,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>363</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline3Cost = 363,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>372</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline4Work = 372,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>373</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline4Cost = 373,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>382</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline5Work = 382,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>383</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline5Cost = 383,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>392</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline6Work = 392,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>393</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline6Cost = 393,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>402</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline7Work = 402,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>403</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline7Cost = 403,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>412</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline8Work = 412,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>413</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline8Cost = 413,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>422</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline9Work = 422,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>423</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline9Cost = 423,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>432</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline10Work = 432,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>433</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline10Cost = 433,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>720</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledActualWorkProtected = 720,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>721</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledActualOvertimeWorkProtected = 721,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>754</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBudgetWork = 754,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>755</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBudgetCost = 755,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>757</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaselineBudgetWork = 757,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>758</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaselineBudgetCost = 758,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>761</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline1BudgetWork = 761,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>762</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline1BudgetCost = 762,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>765</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline2BudgetWork = 765,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>766</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline2BudgetCost = 766,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>769</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline3BudgetWork = 769,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>770</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline3BudgetCost = 770,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>773</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline4BudgetWork = 773,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>774</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline4BudgetCost = 774,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>777</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline5BudgetWork = 777,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>778</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline5BudgetCost = 778,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>781</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline6BudgetWork = 781,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>782</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline6BudgetCost = 782,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>785</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline7BudgetWork = 785,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>786</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline7BudgetCost = 786,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>789</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline8BudgetWork = 789,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>790</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline8BudgetCost = 790,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>793</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline9BudgetWork = 793,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>794</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline9BudgetCost = 794,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>797</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline10BudgetWork = 797,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>798</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14)]
		 pjResourceTimescaledBaseline10BudgetCost = 798,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>811</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledCumulativeActualWork = 811,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>812</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledRemainingCumulativeActualWork = 812,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>813</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledRemainingCumulativeWork = 813,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>814</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaselineRemainingCumulativeWork = 814,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>815</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline1RemainingCumulativeWork = 815,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>816</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline2RemainingCumulativeWork = 816,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>817</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline3RemainingCumulativeWork = 817,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>818</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline4RemainingCumulativeWork = 818,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>819</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline5RemainingCumulativeWork = 819,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>820</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline6RemainingCumulativeWork = 820,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>821</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline7RemainingCumulativeWork = 821,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>822</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline8RemainingCumulativeWork = 822,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>823</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline9RemainingCumulativeWork = 823,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>824</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline10RemainingCumulativeWork = 824,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>825</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaselineCumulativeWork = 825,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>826</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline1CumulativeWork = 826,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>827</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline2CumulativeWork = 827,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>828</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline3CumulativeWork = 828,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>829</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline4CumulativeWork = 829,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>830</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline5CumulativeWork = 830,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>831</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline6CumulativeWork = 831,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>832</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline7CumulativeWork = 832,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>833</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline8CumulativeWork = 833,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>834</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline9CumulativeWork = 834,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>835</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjResourceTimescaledBaseline10CumulativeWork = 835
	}
}