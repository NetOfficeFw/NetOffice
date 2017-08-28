using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsEnum)]
	public enum PivotDataReasonEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonInsertFieldSet = 0,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonRemoveFieldSet = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonInsertTotal = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonRemoveTotal = 3,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAllowDetailsChange = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonSortDirectionChange = 5,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonSortOnChange = 6,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonSortOnScopeChange = 7,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterFunctionChange = 8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterContextChange = 9,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonDisplayCalculatedMembersChange = 10,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterOnChange = 11,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterOnScopeChange = 12,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterFunctionValueChange = 13,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalNameChange = 14,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonIncludedMembersChange = 15,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonExcludedMembersChange = 16,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonIsIncludedChange = 17,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonDisplayEmptyMembersChange = 19,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalFunctionChange = 20,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonUser = 21,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonDataSourceChange = 22,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonDataMemberChange = 23,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonGroupOnChange = 24,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonUnknown = 25,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonGroupStartChange = 26,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonGroupIntervalChange = 27,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonIsFilteredChange = 28,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonOrderedMembersChange = 29,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonGroupEndChange = 30,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonCommandTextChange = 31,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonConnectionStringChange = 32,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonMemberPropertyIsIncludedChange = 33,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonMemberPropertyDisplayInChange = 34,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonSubtotalsChange = 35,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalExpressionChange = 36,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalSolveOrderChange = 37,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalDeleted = 38,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFieldSetDeleted = 39,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonRecordChanged = 40,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAllowMultiFilterChange = 41,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAllIncludeExcludeChange = 42,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAdhocFieldAdded = 43,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAdhocFieldDeleted = 44,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAdhocMemberChanged = 45,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonAlwaysIncludeInCubeChange = 46,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonExpressionChange = 47,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonTotalAllMembersChange = 48,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonDisplayCellColorChange = 49,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFilterCrossJoinsChange = 50,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonRefreshDataSource = 51,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFieldSetNameChange = 52,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("OWC10", 1)]
		 plDataReasonFieldNameChange = 53
	}
}