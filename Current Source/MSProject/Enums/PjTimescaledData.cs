using System;
using LateBindingApi.Core;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjTimescaledData
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjWork = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCost = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCumulativeWork = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCumulativeCost = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualWork = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualOvertimeWork = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjOvertimeWork = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaselineWork = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjOverallocation = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjFixedCost = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualCost = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaselineCost = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjRegularWork = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBCWS = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBCWP = 14,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjACWP = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjSV = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCV = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPercentAllocation = 18,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPeakUnits = 19,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjRemainingAvailability = 20,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjPctComplete = 21,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCumPctComplete = 22,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjAllTaskRows = 23,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjAllResourceRows = 23,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjAllAssignmentRows = 24,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjUnitAvailability = 25,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjWorkAvailability = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline1Work = 27,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline1Cost = 28,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline2Work = 29,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline2Cost = 30,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline3Work = 31,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline3Cost = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline4Work = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline4Cost = 34,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline5Work = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline5Cost = 36,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline6Work = 37,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline6Cost = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline7Work = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline7Cost = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline8Work = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline8Cost = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline9Work = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline9Cost = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline10Work = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline10Cost = 46,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualFixedCost = 47,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCPI = 48,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjSPI = 49,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjCVP = 50,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjSVP = 51,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualWorkProtected = 52,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjActualOvertimeWorkProtected = 53,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBudgetWork = 54,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBudgetCost = 55,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaselineBudgetWork = 56,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaselineBudgetCost = 57,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline1BudgetWork = 58,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline1BudgetCost = 59,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline2BudgetWork = 60,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline2BudgetCost = 61,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline3BudgetWork = 62,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline3BudgetCost = 63,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline4BudgetWork = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline4BudgetCost = 65,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline5BudgetWork = 66,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline5BudgetCost = 67,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline6BudgetWork = 68,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline6BudgetCost = 69,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline7BudgetWork = 70,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline7BudgetCost = 71,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline8BudgetWork = 72,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline8BudgetCost = 73,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline9BudgetWork = 74,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline9BudgetCost = 75,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline10BudgetWork = 76,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjBaseline10BudgetCost = 77
	}
}