using System;
using NetOffice;
namespace NetOffice.PublisherApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Publisher 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PbWebControlType
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlCheckBox = 100,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlCommandButton = 101,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlListBox = 102,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlMultiLineTextBox = 103,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlOptionButton = 104,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlSingleLineTextBox = 105,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlWebComponent = 106,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlHTMLFragment = 107,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWebControlHotSpot = 110
	}
}