using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821354.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcCommand
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWindowUnhide = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWindowHide = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdExit = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCompactDatabase = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEncryptDecryptDatabase = 5,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRepairDatabase = 6,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMakeMDEFile = 7,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMoreWindows = 8,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAppRestore = 9,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAppMaximize = 10,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAppMinimize = 11,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAppMove = 12,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAppSize = 13,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocRestore = 14,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocMaximize = 15,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocMove = 16,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocSize = 17,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRefresh = 18,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFont = 19,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSave = 20,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAs = 21,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWindowCascade = 22,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTileVertically = 23,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWindowArrangeIcons = 24,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenDatabase = 25,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewDatabase = 26,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOLEDDELinks = 27,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecordsGoToNew = 28,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdReplace = 29,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFind = 30,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRunMacro = 31,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPageSetup = 32,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertObject = 33,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDuplicate = 34,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAboutMicrosoftAccess = 35,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFormHdrFtr = 36,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdReportHdrFtr = 37,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPasteAppend = 38,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertFile = 39,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectForm = 40,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTabOrder = 41,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFieldList = 42,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignLeft = 43,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignRight = 44,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignTop = 45,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignBottom = 46,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignToGrid = 47,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSizeToGrid = 48,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOptions = 49,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectRecord = 50,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSortingAndGrouping = 51,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdBringToFront = 52,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSendToBack = 53,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPrintPreview = 54,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdApplyDefault = 55,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSetControlDefaults = 56,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOLEObjectDefaultVerb = 57,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClose = 58,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSizeToFit = 59,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocMinimize = 60,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewRuler = 61,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSnapToGrid = 62,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewGrid = 63,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPasteSpecial = 64,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecordsGoToNext = 65,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecordsGoToPrevious = 66,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecordsGoToFirst = 67,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecordsGoToLast = 68,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSizeToFitForm = 69,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEditingAllowed = 70,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClearGrid = 71,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdJoinProperties = 72,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTotals = 73,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeCrosstab = 74,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTableNames = 75,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryParameters = 76,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFormatCells = 77,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDataEntry = 78,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHideColumns = 79,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUnhideColumns = 80,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteQueryColumn = 81,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertQueryColumn = 82,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRemoveTable = 84,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewToolbox = 85,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMacroNames = 86,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMacroConditions = 87,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSingleStep = 88,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeSelect = 89,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeUpdate = 90,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeAppend = 91,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeDelete = 92,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdApplyFilterSort = 93,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeMakeTable = 94,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdLoadFromQuery = 95,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsQuery = 96,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveRecord = 97,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAdvancedFilterSort = 99,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMicrosoftAccessHelpTopics = 100,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdLinkTables = 102,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUserAndGroupPermissions = 103,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUserAndGroupAccounts = 104,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFreezeColumn = 105,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUnfreezeAllColumns = 106,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPrimaryKey = 107,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSubformDatasheet = 108,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectAllRecords = 109,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTables = 110,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewQueries = 111,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewForms = 112,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewReports = 113,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewMacros = 114,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewModules = 115,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRowHeight = 116,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdColumnWidth = 117,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertFileIntoModule = 118,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveModuleAsText = 119,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFindPrevious = 120,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>121</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWindowSplit = 121,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdProcedureDefinition = 122,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDebugWindow = 123,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdReset = 124,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCompileAllModules = 125,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCompileAndSaveAllModules = 126,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdGoContinue = 127,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStepOver = 128,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSetNextStatement = 129,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowNextStatement = 130,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdToggleBreakpoint = 131,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClearAllBreakpoints = 132,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRelationships = 133,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectTable = 134,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectQuery = 135,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectForm = 136,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectReport = 137,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectMacro = 138,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectModule = 139,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectClassModule = 140,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdLayoutPreview = 141,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsReport = 142,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>143</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRename = 143,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRemoveFilterSort = 144,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>145</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveLayout = 145,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>146</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClearAll = 146,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>147</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHideTable = 147,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>148</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowDirectRelationships = 148,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>149</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowAllRelationships = 149,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateRelationship = 150,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEditRelationship = 151,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdIndexes = 152,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignToShortest = 153,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>154</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAlignToTallest = 154,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>155</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSizeToNarrowest = 155,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>156</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSizeToWidest = 156,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>157</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHorizontalSpacingMakeEqual = 157,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>158</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHorizontalSpacingDecrease = 158,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHorizontalSpacingIncrease = 159,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdVerticalSpacingMakeEqual = 160,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>161</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdVerticalSpacingDecrease = 161,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdVerticalSpacingIncrease = 162,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSortAscending = 163,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSortDescending = 164,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdToolbarsCustomize = 165,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOLEObjectConvert = 167,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeSQLDataDefinition = 168,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeSQLPassThrough = 169,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewCode = 170,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdConvertDatabase = 171,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCallStack = 172,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSend = 173,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOutputToExcel = 175,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOutputToRTF = 176,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOutputToText = 177,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInvokeBuilder = 178,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoomBox = 179,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryTypeSQLUnion = 180,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRun = 181,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPageHdrFtr = 182,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>183</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDesignView = 183,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>184</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSQLView = 184,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>185</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowTable = 185,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>186</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCloseWindow = 186,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>187</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertRows = 187,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>188</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteRows = 188,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>189</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCut = 189,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>190</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCopy = 190,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>191</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPaste = 191,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>192</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAutoDial = 192,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>193</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectAutoForm = 193,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>194</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectAutoReport = 194,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>195</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdWordMailMerge = 195,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>196</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTestValidationRules = 196,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>197</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdControlWizardsToggle = 197,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>198</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEnd = 198,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>199</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRedo = 199,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjectBrowser = 200,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAddWatch = 201,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEditWatch = 202,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQuickWatch = 203,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStepToCursor = 204,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdIndent = 205,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOutdent = 206,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>207</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFilterByForm = 207,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFilterBySelection = 208,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewLargeIcons = 209,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>210</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewDetails = 210,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewSmallIcons = 211,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>212</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewList = 212,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdLineUpIcons = 213,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>214</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdArrangeIconsByName = 214,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>215</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdArrangeIconsByType = 215,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>216</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdArrangeIconsByCreated = 216,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>217</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdArrangeIconsByModified = 217,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>218</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdArrangeIconsAuto = 218,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>219</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateShortcut = 219,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>220</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdToggleFilter = 220,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>221</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenTable = 221,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>222</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertPicture = 222,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>223</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteRecord = 223,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStartupProperties = 224,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>225</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPageNumber = 225,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>226</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDateAndTime = 226,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>227</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToTextBox = 227,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>228</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToLabel = 228,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>229</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToListBox = 229,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>230</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToComboBox = 230,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>231</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToCheckBox = 231,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>232</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToToggleButton = 232,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>233</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToOptionButton = 233,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>234</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdChangeToImage = 234,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>235</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAnswerWizard = 235,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>236</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMicrosoftOnTheWeb = 236,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>237</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClearItemDefaults = 237,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>238</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom200 = 238,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>239</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom150 = 239,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom100 = 240,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>241</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom75 = 241,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>242</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom50 = 242,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>243</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom25 = 243,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>244</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoom10 = 244,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>245</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFitToWindow = 245,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>246</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPreviewOnePage = 246,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>247</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPreviewTwoPages = 247,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>248</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPreviewFourPages = 248,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>249</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPreviewEightPages = 249,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>250</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPreviewTwelvePages = 250,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>251</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenURL = 251,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>252</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenStartPage = 252,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>253</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenSearchPage = 253,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>254</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRegisterActiveXControls = 254,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>255</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteTab = 255,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDatabaseProperties = 256,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>257</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdImport = 257,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>258</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertActiveXControl = 258,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>259</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertHyperlink = 259,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>260</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdReferences = 260,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>261</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAutoCorrect = 261,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>262</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertProcedure = 262,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>263</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateReplica = 263,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>264</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSynchronizeNow = 264,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>265</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRecoverDesignMaster = 265,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>266</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdResolveConflicts = 266,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>267</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteWatch = 267,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>269</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSpelling = 269,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>270</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAutoFormat = 270,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>271</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeleteTableColumn = 271,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>272</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertTableColumn = 272,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>273</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertLookupColumn = 273,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>274</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRenameColumn = 274,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>275</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSetDatabasePassword = 275,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>276</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUserLevelSecurityWizard = 276,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>277</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFilterExcludingSelection = 277,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>278</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQuickPrint = 278,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>279</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdConvertMacrosToVisualBasic = 279,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>280</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAllModules = 280,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>281</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFormView = 281,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>282</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDatasheetView = 282,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>283</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAnalyzePerformance = 283,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>284</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdAnalyzeTable = 284,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>285</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDocumenter = 285,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>286</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTileHorizontally = 286,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>287</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdProperties = 287,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>288</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTransparentBackground = 288,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>289</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTransparentBorder = 289,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>290</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCompileLoadedModules = 290,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>291</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertLookupField = 291,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>292</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdUndo = 292,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>293</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertChart = 293,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>294</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdGoBack = 294,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>295</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdGoForward = 295,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>296</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStopLoadingPage = 296,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>297</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRefreshPage = 297,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>298</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFavoritesOpen = 298,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>299</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFavoritesAddTo = 299,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowOnlyWebToolbar = 300,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>301</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdToolbarControlProperties = 301,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>302</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdShowMembers = 302,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>303</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdListConstants = 303,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>304</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQuickInfo = 304,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>305</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdParameterInfo = 305,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>306</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCompleteWord = 306,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>307</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdBookmarksToggle = 307,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>308</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdBookmarksNext = 308,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>309</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdBookmarksPrevious = 309,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>310</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdBookmarksClearAll = 310,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>311</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStepOut = 311,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>312</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFindPrevWordUnderCursor = 312,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>313</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFindNextWordUnderCursor = 313,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>314</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjBrwFindWholeWordOnly = 314,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>315</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjBrwShowHiddenMembers = 315,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>316</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjBrwHelp = 316,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>317</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjBrwViewDefinition = 317,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>318</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdObjBrwGroupMembers = 318,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>319</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectReport = 319,

		 /// <summary>
		 /// SupportByVersion Access 9, 10
		 /// </summary>
		 /// <remarks>320</remarks>
		 [SupportByVersion("Access", 9,10)]
		 acCmdPublish = 320,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>321</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsHTML = 321,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>322</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsIDC = 322,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>323</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsASP = 323,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>324</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPublishDefaults = 324,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>325</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEditHyperlink = 325,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>326</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenHyperlink = 326,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>327</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdOpenNewHyperlink = 327,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>328</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCopyHyperlink = 328,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>329</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHyperlinkDisplayText = 329,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>330</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdTabControlPageOrder = 330,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>331</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdInsertPage = 331,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>332</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDeletePage = 332,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>333</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectAll = 333,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>334</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateMenuFromMacro = 334,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>335</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateToolbarFromMacro = 335,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>336</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdCreateShortcutMenuFromMacro = 336,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>337</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDelete = 337,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>338</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRunOpenMacro = 338,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>339</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdLastPosition = 339,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>340</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPrint = 340,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>341</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdFindNext = 341,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>342</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdStepInto = 342,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>343</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdClearHyperlink = 343,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>344</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDataAccessPageBrowse = 344,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>346</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectDataAccessPage = 346,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>347</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSelectDataAccessPage = 347,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>349</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewDataAccessPages = 349,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>350</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectView = 350,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>351</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectStoredProcedure = 351,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>352</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdNewObjectDiagram = 352,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewFieldList = 353,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>354</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewDiagrams = 354,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>355</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewStoredProcedures = 355,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>356</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewViews = 356,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>357</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewShowPaneSQL = 357,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>358</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewShowPaneDiagram = 358,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>359</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewShowPaneGrid = 359,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>360</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewVerifySQL = 360,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>361</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryGroupBy = 361,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>362</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdQueryAddToOutput = 362,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>363</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTableColumnNames = 363,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>364</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTableNameOnly = 364,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>365</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdHidePane = 365,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>366</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdRemove = 366,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>368</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTableColumnProperties = 368,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>369</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTableKeys = 369,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>370</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdViewTableUserView = 370,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>371</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdZoomSelection = 371,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>372</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramNewLabel = 372,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>373</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramAddRelatedTables = 373,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>374</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramShowRelationshipLabels = 374,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>375</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramModifyUserDefinedView = 375,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>376</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramViewPageBreaks = 376,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>377</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramRecalculatePageBreaks = 377,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>378</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramAutosizeSelectedTables = 378,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>379</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramLayoutSelection = 379,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>380</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramLayoutDiagram = 380,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>381</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramNewTable = 381,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>382</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDiagramDeleteRelationship = 382,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>383</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdConnection = 383,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>384</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdEditTriggers = 384,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>385</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDataAccessPageDesignView = 385,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>386</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdPromote = 386,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>387</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdGroupByTable = 387,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>388</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdDemote = 388,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>389</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdSaveAsDataAccessPage = 389,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>390</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acCmdMicrosoftScriptEditor = 390,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>391</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdWorkgroupAdministrator = 391,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>394</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdNewObjectFunction = 394,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>395</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdViewFunctions = 395,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>396</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableView = 396,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>397</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartView = 397,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>398</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoFilter = 398,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>399</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableSubtotal = 399,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotCollapse = 400,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>401</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotExpand = 401,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>402</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableHideDetails = 402,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>403</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableShowDetails = 403,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>404</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotRefresh = 404,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>405</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableExportToExcel = 405,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>406</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableMoveToRowArea = 406,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>407</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableMoveToColumnArea = 407,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>408</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableMoveToFilterArea = 408,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>409</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableMoveToDetailArea = 409,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>410</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePromote = 410,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>411</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableDemote = 411,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>412</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoSum = 412,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>413</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoCount = 413,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>414</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoMin = 414,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>415</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoMax = 415,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>416</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoAverage = 416,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>417</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoStdDev = 417,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>418</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoVar = 418,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>419</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoStdDevP = 419,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>420</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotAutoVarP = 420,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>421</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop1 = 421,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>422</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop2 = 422,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>423</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop5 = 423,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>424</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop10 = 424,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>425</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop25 = 425,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>426</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop1Percent = 426,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>427</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop2Percent = 427,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>428</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop5Percent = 428,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>429</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop10Percent = 429,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>430</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTop25Percent = 430,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>431</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowTopOther = 431,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>432</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom1 = 432,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>433</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom2 = 433,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>434</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom5 = 434,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>435</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom10 = 435,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>436</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom25 = 436,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>437</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom1Percent = 437,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>438</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom2Percent = 438,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>439</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom5Percent = 439,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>440</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom10Percent = 440,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>441</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottom25Percent = 441,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>442</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowBottomOther = 442,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>443</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableCreateCalcTotal = 443,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>444</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableCreateCalcField = 444,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>445</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableShowAsNormal = 445,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>446</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePercentRowTotal = 446,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>447</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePercentColumnTotal = 447,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>448</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePercentParentRowItem = 448,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>449</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePercentParentColumnItem = 449,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>450</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTablePercentGrandTotal = 450,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>451</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableExpandIndicators = 451,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>452</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotDropAreas = 452,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>453</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartType = 453,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>454</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotDelete = 454,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>455</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartShowLegend = 455,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>456</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartByRowByColumn = 456,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>457</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartDrillInto = 457,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>458</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartMultiplePlots = 458,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>459</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartMultiplePlotsUnifiedScale = 459,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>460</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartUndo = 460,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>461</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotShowAll = 461,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>462</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubformFormView = 462,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>463</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubformDatasheetView = 463,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>464</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubformPivotTableView = 464,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>465</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubformPivotChartView = 465,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>466</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdWebPagePreview = 466,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>467</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPageProperties = 467,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>468</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDataOutline = 468,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>469</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdInsertMovieFromFile = 469,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>470</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdInsertPivotTable = 470,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>471</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdInsertSpreadsheet = 471,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>472</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdInsertUnboundSection = 472,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>473</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdWebTheme = 473,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>474</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdBackgroundPicture = 474,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>475</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdBackgroundSound = 475,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>476</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdAlignMiddle = 476,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>477</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdAlignCenter = 477,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>478</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdAlignmentAndSizing = 478,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>479</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDataAccessPageFieldListRefresh = 479,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>480</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDataAccessPageAddToPage = 480,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>481</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdZoom500 = 481,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>482</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdZoom1000 = 482,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>483</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPrintRelationships = 483,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>484</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdGroupControls = 484,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>485</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdUngroupControls = 485,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>486</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdWebPageProperties = 486,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>487</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdExport = 487,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>488</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdOfficeClipboard = 488,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>489</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDeleteTable = 489,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>490</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPasteAsHyperlink = 490,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>491</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdNewGroup = 491,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>492</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdRenameGroup = 492,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>493</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDeleteGroup = 493,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>494</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdAddToNewGroup = 494,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>495</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubformInNewWindow = 495,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>496</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdServerProperties = 496,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>497</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdTableCustomView = 497,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>498</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdTableAddTable = 498,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>499</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdInsertSubdatasheet = 499,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdConditionalFormatting = 500,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdChangeToCommandButton = 501,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>504</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubdatasheetExpandAll = 504,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubdatasheetCollapseAll = 505,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSubdatasheetRemove = 506,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>507</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdServerFilterByForm = 507,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>508</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdMaximiumRecords = 508,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>511</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSpeech = 511,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>513</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdBackup = 513,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>514</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdRestore = 514,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>515</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdTransferSQLDatabase = 515,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>516</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdCopyDatabaseFile = 516,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>517</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDropSQLDatabase = 517,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>519</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdLinkedTableManager = 519,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>520</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdDatabaseSplitter = 520,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>521</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdSwitchboardManager = 521,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>522</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdUpsizingWizard = 522,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>524</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPartialReplicaWizard = 524,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>525</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdVisualBasicEditor = 525,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>526</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdAddInManager = 526,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>527</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableClearCustomOrdering = 527,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>528</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableFilterBySelection = 528,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>529</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableRemove = 529,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>530</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableGroupItems = 530,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>531</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotTableUngroupItems = 531,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>532</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartDrillOut = 532,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>533</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdShowEnvelope = 533,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>534</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartSortAscByTotal = 534,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>535</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acCmdPivotChartSortDescByTotal = 535,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>536</remarks>
		 [SupportByVersion("Access", 11,12,14,15,16)]
		 acCmdViewObjectDependencies = 536,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>537</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdPublishDatabase = 537,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>538</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdCloseDatabase = 538,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>539</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdReportView = 539,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>540</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdToggleOffline = 540,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>541</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRefreshData = 541,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>542</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdShareOnSharePoint = 542,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>543</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSavedImports = 543,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>544</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachAccess = 544,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>545</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachExcel = 545,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>546</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachText = 546,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>547</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachSharePointList = 547,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>548</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachXML = 548,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>549</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachODBC = 549,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>550</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachHTML = 550,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>551</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachOutlook = 551,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>552</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachdBase = 552,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>553</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachParadox = 553,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>554</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdImportAttachLotus = 554,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>555</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSavedExports = 555,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>556</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportExcel = 556,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>557</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportSharePointList = 557,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>558</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportRTF = 558,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>559</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportAccess = 559,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>560</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportText = 560,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>561</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportXML = 561,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>562</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportODBC = 562,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>563</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportSnapShot = 563,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>564</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportHTML = 564,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>565</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportdBase = 565,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>566</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportParadox = 566,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>567</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportLotus = 567,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>568</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdStackedLayout = 568,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>569</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdTabularLayout = 569,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>570</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSelectEntireRow = 570,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>571</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSelectEntireColumn = 571,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>572</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdMoveColumnCellUp = 572,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>573</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdMoveColumnCellDown = 573,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>574</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesBoth = 574,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>575</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesVertical = 575,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>576</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesHorizontal = 576,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>577</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesNone = 577,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>578</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesCrossHatch = 578,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>579</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesTop = 579,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>580</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesBottom = 580,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>581</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutGridlinesOutline = 581,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>582</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRemoveFromLayout = 582,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>583</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAddFromOutlook = 583,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>584</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSaveAsOutlookContact = 584,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>585</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdInsertLogo = 585,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>586</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdInsertTitle = 586,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>587</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdPasteFormatting = 587,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>588</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdMacroArguments = 588,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>589</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdMacroAllActions = 589,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>590</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdPrintSelection = 590,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>591</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdPublishFixedFormat = 591,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>592</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdExportFixedFormat = 592,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>593</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdLayoutView = 593,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>594</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectContinuousForm = 594,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>595</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectSplitForm = 595,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>596</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectPivotChart = 596,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>597</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectPivotTable = 597,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>598</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectDatasheetForm = 598,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>599</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectModalForm = 599,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>600</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectBlankForm = 600,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>601</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectLabelsReport = 601,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>602</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectBlankReport = 602,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>603</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectDesignQuery = 603,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>604</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectDesignForm = 604,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>605</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectDesignReport = 605,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>606</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdNewObjectDesignTable = 606,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>607</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdEditListItems = 607,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>608</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdCollectDataViaEmail = 608,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>609</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdManageReplies = 609,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>610</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorTopLeft = 610,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>611</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorStretchAcross = 611,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>612</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorTopRight = 612,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>613</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorStretchDown = 613,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>614</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorStretchDownAcross = 614,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>615</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorStretchDownRight = 615,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>616</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorBottomLeft = 616,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>617</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorBottomStretchAcross = 617,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>618</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdAnchorBottomRight = 618,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>619</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdFilterMenu = 619,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>620</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdShowColumnHistory = 620,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>621</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdBrowseSharePointList = 621,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>622</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdModifySharePointList = 622,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>623</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdModifySharePointListAlerts = 623,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>624</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdModifySharePointListWorkflow = 624,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>625</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdModifySharePointListPermissions = 625,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>626</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRefreshSharePointList = 626,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>627</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdDeleteSharePointList = 627,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>628</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlMarginsNone = 628,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>629</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlMarginsNarrow = 629,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>630</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlMarginsMedium = 630,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>631</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlMarginsWide = 631,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>632</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlPaddingNone = 632,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>633</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlPaddingNarrow = 633,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>634</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlPaddingMedium = 634,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>635</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdControlPaddingWide = 635,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>636</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdShowDatePicker = 636,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>637</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSetCaption = 637,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>638</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSynchronize = 638,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>639</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdDiscardChanges = 639,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>640</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdDiscardChangesAndRefresh = 640,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>641</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdSharePointSiteRecycleBin = 641,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>642</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdToggleCacheListData = 642,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>643</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRemoveFilterFromCurrentColumn = 643,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>644</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRemoveAllFilters = 644,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>645</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdRemoveAllSorts = 645,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>646</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdCloseAll = 646,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>647</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdFieldTemplates = 647,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>648</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat1 = 648,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>649</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat2 = 649,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>650</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat3 = 650,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>651</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat4 = 651,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>652</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat5 = 652,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>653</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat6 = 653,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>654</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat7 = 654,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>655</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat8 = 655,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>656</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat9 = 656,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>657</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat10 = 657,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>658</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat11 = 658,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>659</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat12 = 659,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>660</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat13 = 660,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>661</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat14 = 661,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>662</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat15 = 662,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>663</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat16 = 663,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>664</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat17 = 664,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>665</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat18 = 665,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>666</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat19 = 666,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>667</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat20 = 667,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>668</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat21 = 668,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>669</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat22 = 669,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>670</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat23 = 670,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>671</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat24 = 671,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>672</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdApplyAutoFormat25 = 672,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>673</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdManageAttachments = 673,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>674</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdWorkflowTasks = 674,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>675</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdStartNewWorkflow = 675,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>676</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdShowMessageBar = 676,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>677</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acCmdHideMessageBar = 677,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>678</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutInsertRowAbove = 678,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>679</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutInsertRowBelow = 679,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>680</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutInsertColumnLeft = 680,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>681</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutInsertColumnRight = 681,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>682</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutMergeCells = 682,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>683</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutSplitColumnCell = 683,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>684</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdLayoutSplitRowCell = 684,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>685</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdShowLogicCatalog = 685,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>686</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdSaveAsTemplate = 686,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>687</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdSaveDatabaseAsNewTemplatePart = 687,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>688</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdSaveSelectionAsNewDataType = 688,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>689</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTop = 689,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>690</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationLeft = 690,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>691</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationRight = 691,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>692</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopTop = 692,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>693</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopLeft = 693,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>694</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopRight = 694,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>695</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdCompatCheckDatabase = 695,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>696</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdCompatCheckCurrentObject = 696,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>697</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdDesignObject = 697,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>698</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdCalculatedColumn = 698,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>699</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdSyncWebApplication = 699,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>700</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdConvertLinkedTableToLocal = 700,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectContinuousFormWeb = 701,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectDatasheetFormWeb = 702,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>703</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectBlankFormWeb = 703,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>704</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectBlankReportWeb = 704,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>705</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectAutoFormWeb = 705,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>706</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectAutoReportWeb = 706,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>707</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectQueryWeb = 707,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>708</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectMacroWeb = 708,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>709</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopWeb = 709,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>710</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationLeftWeb = 710,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>711</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationRightWeb = 711,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>712</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopTopWeb = 712,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>713</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopLeftWeb = 713,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>714</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdNewObjectNavigationTopRightWeb = 714,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>715</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdSelectEntireLayout = 715,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>716</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdPrepareDatabaseForWeb = 716,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>717</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdManageTableEvents = 717,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>718</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddDataMacroAfterInsert = 718,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>719</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddDataMacroAfterUpdate = 719,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>720</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddDataMacroAfterDelete = 720,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>721</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddDataMacroBeforeDelete = 721,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>722</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddDataMacroBeforeChange = 722,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>723</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdAddNamedDataMacro = 723,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>724</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acCmdInsertNavigationButton = 724
	}
}