using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlObjectClass
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olApplication = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olNamespace = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olFolder = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olRecipient = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAttachment = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAddressList = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAddressEntry = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olFolders = 15,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olItems = 16,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olRecipients = 17,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAttachments = 18,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAddressLists = 20,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAddressEntries = 21,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAppointment = 26,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMeetingRequest = 53,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMeetingCancellation = 54,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMeetingResponseNegative = 55,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMeetingResponsePositive = 56,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMeetingResponseTentative = 57,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olRecurrencePattern = 28,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olExceptions = 29,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olException = 30,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olAction = 32,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olActions = 33,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olExplorer = 34,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olInspector = 35,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPages = 36,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olFormDescription = 37,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olUserProperties = 38,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olUserProperty = 39,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olContact = 40,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olDocument = 41,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olJournal = 42,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olMail = 43,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olNote = 44,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPost = 45,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olReport = 46,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olRemote = 47,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTask = 48,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTaskRequest = 49,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTaskRequestUpdate = 50,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTaskRequestAccept = 51,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olTaskRequestDecline = 52,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olExplorers = 60,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olInspectors = 61,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPanes = 62,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarPane = 63,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarStorage = 64,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarGroups = 65,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarGroup = 66,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarShortcuts = 67,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olOutlookBarShortcut = 68,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olDistributionList = 69,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPropertyPageSite = 70,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olPropertyPages = 71,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olSyncObject = 72,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olSyncObjects = 73,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olSelection = 74,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olLink = 75,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersionAttribute("Outlook", 9,10,11,12,14)]
		 olLinks = 76,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olSearch = 77,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olResults = 78,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olViews = 79,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olView = 80,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olItemProperties = 98,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olItemProperty = 99,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olReminders = 100,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("Outlook", 10,11,12,14)]
		 olReminder = 101,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olConflict = 102,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Outlook", 11,12,14)]
		 olConflicts = 103,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olSharing = 104,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAccount = 105,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAccounts = 106,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olStore = 107,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olStores = 108,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olSelectNamesDialog = 109,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeUser = 110,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olExchangeDistributionList = 111,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olPropertyAccessor = 112,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olStorageItem = 113,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRules = 114,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRule = 115,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRuleActions = 116,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRuleAction = 117,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMoveOrCopyRuleAction = 118,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olSendRuleAction = 119,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olTable = 120,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>121</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRow = 121,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAssignToCategoryRuleAction = 122,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olPlaySoundRuleAction = 123,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMarkAsTaskRuleAction = 124,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNewItemAlertRuleAction = 125,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRuleConditions = 126,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olRuleCondition = 127,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olImportanceRuleCondition = 128,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormRegion = 129,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olCategoryRuleCondition = 130,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormNameRuleCondition = 131,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFromRuleCondition = 132,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olSenderInAddressListRuleCondition = 133,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olTextRuleCondition = 134,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAccountRuleCondition = 135,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassTableView = 136,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassIconView = 137,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassCardView = 138,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassCalendarView = 139,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassTimeLineView = 140,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olViewFields = 141,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olViewField = 142,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olOrderField = 144,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>145</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olOrderFields = 145,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>146</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olViewFont = 146,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>147</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAutoFormatRule = 147,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>148</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAutoFormatRules = 148,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>149</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olColumnFormat = 149,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olColumns = 150,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olCalendarSharing = 151,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olCategory = 152,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olCategories = 153,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>154</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olColumn = 154,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>155</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassNavigationPane = 155,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>156</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationModules = 156,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>157</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationModule = 157,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>158</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMailModule = 158,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olCalendarModule = 159,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olContactsModule = 160,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>161</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olTasksModule = 161,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olJournalModule = 162,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNotesModule = 163,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationGroups = 164,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationGroup = 165,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationFolders = 166,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olNavigationFolder = 167,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassBusinessCardView = 168,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAttachmentSelection = 169,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olAddressRuleCondition = 170,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olUserDefinedProperty = 171,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olUserDefinedProperties = 172,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFromRssFeedRuleCondition = 173,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>174</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassTimeZone = 174,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olClassTimeZones = 175,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olMobile = 176,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olSolutionsModule = 177,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olConversation = 178,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olSimpleItems = 179,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olOutspace = 180,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olMeetingForwardNotification = 181,

		 /// <summary>
		 /// SupportByVersion Outlook 14
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersionAttribute("Outlook", 14)]
		 olConversationHeader = 182
	}
}