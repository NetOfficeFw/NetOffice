using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866800(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjAssignmentTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledWork = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledRegularWork = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledOvertimeWork = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledActualWork = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledActualOvertimeWork = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledCumulativeWork = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaselineWork = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledOverallocation = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledPercentAllocation = 18,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledPeakUnits = 19,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledCost = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledActualCost = 28,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaselineCost = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledCumulativeCost = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBCWS = 34,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBCWP = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledACWP = 36,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledSV = 37,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>247</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledCV = 247,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>289</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline1Work = 289,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>290</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline1Cost = 290,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>298</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline2Work = 298,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>299</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline2Cost = 299,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>307</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline3Work = 307,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>308</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline3Cost = 308,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>316</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline4Work = 316,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>317</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline4Cost = 317,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>325</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline5Work = 325,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>326</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline5Cost = 326,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>334</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline6Work = 334,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>335</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline6Cost = 335,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>343</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline7Work = 343,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>344</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline7Cost = 344,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>352</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline8Work = 352,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline8Cost = 353,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>361</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline9Work = 361,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>362</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline9Cost = 362,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>370</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline10Work = 370,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>371</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline10Cost = 371,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>630</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledActualWorkProtected = 630,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>631</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledActualOvertimeWorkProtected = 631,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>669</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBudgetWork = 669,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>670</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBudgetCost = 670,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>673</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaselineBudgetWork = 673,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>674</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaselineBudgetCost = 674,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>677</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline1BudgetWork = 677,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>678</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline1BudgetCost = 678,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>681</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline2BudgetWork = 681,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>682</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline2BudgetCost = 682,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>685</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline3BudgetWork = 685,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>686</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline3BudgetCost = 686,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>689</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline4BudgetWork = 689,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>690</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline4BudgetCost = 690,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>693</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline5BudgetWork = 693,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>694</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline5BudgetCost = 694,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>697</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline6BudgetWork = 697,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>698</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline6BudgetCost = 698,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline7BudgetWork = 701,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline7BudgetCost = 702,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>705</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline8BudgetWork = 705,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>706</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline8BudgetCost = 706,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>709</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline9BudgetWork = 709,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>710</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline9BudgetCost = 710,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>713</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline10BudgetWork = 713,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>714</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjAssignmentTimescaledBaseline10BudgetCost = 714,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>727</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledCumulativeActualWork = 727,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>728</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledRemainingCumulativeActualWork = 728,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>729</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledRemainingCumulativeWork = 729,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>730</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaselineRemainingCumulativeWork = 730,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>731</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline1RemainingCumulativeWork = 731,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>732</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline2RemainingCumulativeWork = 732,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>733</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline3RemainingCumulativeWork = 733,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>734</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline4RemainingCumulativeWork = 734,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>735</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline5RemainingCumulativeWork = 735,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>736</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline6RemainingCumulativeWork = 736,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>737</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline7RemainingCumulativeWork = 737,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>738</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline8RemainingCumulativeWork = 738,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>739</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline9RemainingCumulativeWork = 739,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>740</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline10RemainingCumulativeWork = 740,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>752</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaselineCumulativeWork = 752,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>753</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline1CumulativeWork = 753,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>754</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline2CumulativeWork = 754,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>755</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline3CumulativeWork = 755,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>756</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline4CumulativeWork = 756,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>757</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline5CumulativeWork = 757,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>758</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline6CumulativeWork = 758,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>759</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline7CumulativeWork = 759,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>760</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline8CumulativeWork = 760,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>761</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline9CumulativeWork = 761,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>762</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjAssignmentTimescaledBaseline10CumulativeWork = 762
	}
}