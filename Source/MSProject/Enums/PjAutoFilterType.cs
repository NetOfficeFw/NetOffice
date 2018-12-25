using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 14, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867954(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjAutoFilterType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterClear = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterCustom = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterIn = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterToday = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterTomorrow = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterYesterday = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNextWeek = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterThisWeek = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterLastWeek = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNextMonth = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterThisMonth = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterLastMonth = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNextQuarter = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterThisQuarter = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterLastQuarter = 14,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNextYear = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterThisYear = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterLastYear = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterYearToDate = 18,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterBeforeStatusDate = 19,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterAfterStatusDate = 20,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNoDate = 21,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilter1DayOrLess = 22,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterBetween1DayAnd1Week = 23,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilter1WeekOrLonger = 24,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNoDuration = 25,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilter8HoursOrLess = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterBetween8And40Hours = 27,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilter40HoursOrMore = 28,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNoWork = 29,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNotStarted = 30,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterInProgress = 31,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterComplete = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterWithin1And25Percent = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterWithin26And50Percent = 34,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterWithin51And75Percent = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterWithin76And100Percent = 36,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterGreaterThan0Cost = 37,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterFlagYes = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterFlagNo = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterEquals = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNotEquals = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterGreaterThan = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterLessThan = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterBetween = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterBeginsWith = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterEndsWith = 46,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterContains = 47,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterNotContains = 48,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic0 = 49,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic1 = 50,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic2 = 51,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic3 = 52,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic4 = 53,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic5 = 54,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic6 = 55,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic7 = 56,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic8 = 57,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic9 = 58,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 14, 16
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("MSProject", 11,14,16)]
		 pjAutoFilterPhonetic10 = 59
	}
}