using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194205.aspx </remarks>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcControlType
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acLabel = 100,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acRectangle = 101,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acLine = 102,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acImage = 103,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acCommandButton = 104,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acOptionButton = 105,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acCheckBox = 106,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acOptionGroup = 107,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acBoundObjectFrame = 108,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acTextBox = 109,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acListBox = 110,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acComboBox = 111,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acSubform = 112,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acObjectFrame = 114,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acPageBreak = 118,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acCustomControl = 119,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acToggleButton = 122,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acTabCtl = 123,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		 acPage = 124,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15,16)]
		 acAttachment = 126,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersionAttribute("Access", 14,15,16)]
		 acEmptyCell = 127,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Access", 14,15,16)]
		 acWebBrowser = 128,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersionAttribute("Access", 14,15,16)]
		 acNavigationControl = 129,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersionAttribute("Access", 14,15,16)]
		 acNavigationButton = 130
	}
}