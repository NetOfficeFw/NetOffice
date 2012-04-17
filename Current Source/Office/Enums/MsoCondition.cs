using System;
using LateBindingApi.Core;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoCondition
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeAllFiles = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeOfficeFiles = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeWordDocuments = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeExcelWorkbooks = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypePowerPointPresentations = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeBinders = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeDatabases = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionFileTypeTemplates = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIncludes = 9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIncludesPhrase = 10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionBeginsWith = 11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionEndsWith = 12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIncludesNearEachOther = 13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIsExactly = 14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIsNot = 15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionYesterday = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionToday = 17,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionTomorrow = 18,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionLastWeek = 19,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionThisWeek = 20,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionNextWeek = 21,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionLastMonth = 22,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionThisMonth = 23,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionNextMonth = 24,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionAnytime = 25,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionAnytimeBetween = 26,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionOn = 27,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionOnOrAfter = 28,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionOnOrBefore = 29,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionInTheNext = 30,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionInTheLast = 31,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionEquals = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionDoesNotEqual = 33,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionAnyNumberBetween = 34,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionAtMost = 35,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionAtLeast = 36,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionMoreThan = 37,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionLessThan = 38,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIsYes = 39,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoConditionIsNo = 40,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionIncludesFormsOf = 41,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFreeText = 42,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeOutlookItems = 43,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeMailItem = 44,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeCalendarItem = 45,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeContactItem = 46,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeNoteItem = 47,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeJournalItem = 48,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeTaskItem = 49,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypePhotoDrawFiles = 50,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeDataConnectionFiles = 51,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypePublisherFiles = 52,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeProjectFiles = 53,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeDocumentImagingFiles = 54,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeVisioFiles = 55,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeDesignerFiles = 56,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionFileTypeWebPages = 57,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsLow = 58,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsNormal = 59,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsHigh = 60,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToLow = 61,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToNormal = 62,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToHigh = 63,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsNotStarted = 64,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsInProgress = 65,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsCompleted = 66,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsWaitingForSomeoneElse = 67,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionEqualsDeferred = 68,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToNotStarted = 69,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToInProgress = 70,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToCompleted = 71,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToWaitingForSomeoneElse = 72,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoConditionNotEqualToDeferred = 73
	}
}