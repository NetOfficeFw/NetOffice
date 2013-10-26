using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192211.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdFieldType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldEmpty = -1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldRef = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldIndexEntry = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFootnoteRef = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSet = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldIf = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldIndex = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTOCEntry = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldStyleRef = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldRefDoc = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSequence = 12,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTOC = 13,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldInfo = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTitle = 15,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSubject = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAuthor = 17,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldKeyWord = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldComments = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldLastSavedBy = 20,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldCreateDate = 21,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSaveDate = 22,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldPrintDate = 23,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldRevisionNum = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldEditTime = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNumPages = 26,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNumWords = 27,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNumChars = 28,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFileName = 29,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTemplate = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDate = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTime = 32,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldPage = 33,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldExpression = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldQuote = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldInclude = 36,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldPageRef = 37,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAsk = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFillIn = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldData = 40,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNext = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNextIf = 42,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSkipIf = 43,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldMergeRec = 44,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDDE = 45,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDDEAuto = 46,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldGlossary = 47,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldPrint = 48,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFormula = 49,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldGoToButton = 50,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldMacroButton = 51,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAutoNumOutline = 52,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAutoNumLegal = 53,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAutoNum = 54,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldImport = 55,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldLink = 56,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSymbol = 57,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldEmbed = 58,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldMergeField = 59,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldUserName = 60,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldUserInitials = 61,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldUserAddress = 62,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldBarCode = 63,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDocVariable = 64,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSection = 65,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSectionPages = 66,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldIncludePicture = 67,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldIncludeText = 68,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFileSize = 69,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFormTextInput = 70,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFormCheckBox = 71,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldNoteRef = 72,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTOA = 73,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldTOAEntry = 74,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldMergeSeq = 75,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldPrivate = 77,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDatabase = 78,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAutoText = 79,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldCompare = 80,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAddin = 81,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldSubscriber = 82,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldFormDropDown = 83,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAdvance = 84,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldDocProperty = 85,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldOCX = 87,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldHyperlink = 88,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldAutoTextList = 89,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldListNum = 90,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFieldHTMLActiveX = 91,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdFieldBidiOutline = 92,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdFieldAddressBlock = 93,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdFieldGreetingLine = 94,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdFieldShape = 95,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdFieldCitation = 96,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdFieldBibliography = 97,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdFieldMergeBarcode = 98,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdFieldDisplayBarcode = 99
	}
}