using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863329.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlObjectClass
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olApplication = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olNamespace = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olFolder = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olRecipient = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAttachment = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAddressList = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAddressEntry = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olFolders = 15,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olItems = 16,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olRecipients = 17,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAttachments = 18,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAddressLists = 20,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAddressEntries = 21,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAppointment = 26,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMeetingRequest = 53,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMeetingCancellation = 54,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMeetingResponseNegative = 55,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMeetingResponsePositive = 56,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMeetingResponseTentative = 57,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olRecurrencePattern = 28,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olExceptions = 29,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olException = 30,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olAction = 32,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olActions = 33,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olExplorer = 34,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olInspector = 35,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olPages = 36,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olFormDescription = 37,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olUserProperties = 38,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olUserProperty = 39,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olContact = 40,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olDocument = 41,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olJournal = 42,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMail = 43,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olNote = 44,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olPost = 45,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olReport = 46,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olRemote = 47,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTask = 48,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTaskRequest = 49,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTaskRequestUpdate = 50,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTaskRequestAccept = 51,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTaskRequestDecline = 52,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olExplorers = 60,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olInspectors = 61,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olPanes = 62,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarPane = 63,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarStorage = 64,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarGroups = 65,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarGroup = 66,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarShortcuts = 67,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olOutlookBarShortcut = 68,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olDistributionList = 69,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olPropertyPageSite = 70,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olPropertyPages = 71,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olSyncObject = 72,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olSyncObjects = 73,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olSelection = 74,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olLink = 75,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olLinks = 76,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olSearch = 77,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olResults = 78,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olViews = 79,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olView = 80,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olItemProperties = 98,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olItemProperty = 99,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olReminders = 100,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olReminder = 101,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olConflict = 102,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olConflicts = 103,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olSharing = 104,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAccount = 105,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAccounts = 106,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olStore = 107,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olStores = 108,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olSelectNamesDialog = 109,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olExchangeUser = 110,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olExchangeDistributionList = 111,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olPropertyAccessor = 112,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olStorageItem = 113,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRules = 114,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRule = 115,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActions = 116,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleAction = 117,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olMoveOrCopyRuleAction = 118,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olSendRuleAction = 119,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olTable = 120,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>121</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRow = 121,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAssignToCategoryRuleAction = 122,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olPlaySoundRuleAction = 123,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olMarkAsTaskRuleAction = 124,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNewItemAlertRuleAction = 125,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleConditions = 126,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleCondition = 127,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olImportanceRuleCondition = 128,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormRegion = 129,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olCategoryRuleCondition = 130,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFormNameRuleCondition = 131,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFromRuleCondition = 132,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olSenderInAddressListRuleCondition = 133,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olTextRuleCondition = 134,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAccountRuleCondition = 135,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassTableView = 136,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassIconView = 137,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassCardView = 138,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassCalendarView = 139,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassTimeLineView = 140,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olViewFields = 141,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olViewField = 142,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olOrderField = 144,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>145</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olOrderFields = 145,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>146</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olViewFont = 146,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>147</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAutoFormatRule = 147,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>148</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAutoFormatRules = 148,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>149</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olColumnFormat = 149,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olColumns = 150,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olCalendarSharing = 151,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olCategory = 152,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olCategories = 153,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>154</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olColumn = 154,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>155</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassNavigationPane = 155,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>156</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationModules = 156,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>157</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationModule = 157,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>158</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olMailModule = 158,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olCalendarModule = 159,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olContactsModule = 160,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>161</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olTasksModule = 161,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olJournalModule = 162,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNotesModule = 163,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationGroups = 164,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationGroup = 165,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationFolders = 166,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olNavigationFolder = 167,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassBusinessCardView = 168,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAttachmentSelection = 169,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olAddressRuleCondition = 170,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olUserDefinedProperty = 171,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olUserDefinedProperties = 172,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olFromRssFeedRuleCondition = 173,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>174</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassTimeZone = 174,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olClassTimeZones = 175,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olMobile = 176,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olSolutionsModule = 177,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olConversation = 178,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olSimpleItems = 179,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olOutspace = 180,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olMeetingForwardNotification = 181,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15, 16
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersion("Outlook", 14,15,16)]
		 olConversationHeader = 182,

		 /// <summary>
		 /// SupportByVersion Outlook 15,16
		 /// </summary>
		 /// <remarks>183</remarks>
		 [SupportByVersion("Outlook", 15, 16)]
		 olClassPeopleView = 183
	}
}