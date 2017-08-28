using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PublisherApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Publisher 14, 15, 16
	 /// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PbWebControlType
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlCheckBox = 100,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlCommandButton = 101,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlListBox = 102,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlMultiLineTextBox = 103,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlOptionButton = 104,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlSingleLineTextBox = 105,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlWebComponent = 106,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlHTMLFragment = 107,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("Publisher", 14,15,16)]
		 pbWebControlHotSpot = 110
	}
}