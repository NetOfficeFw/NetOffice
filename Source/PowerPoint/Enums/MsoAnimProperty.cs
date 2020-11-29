﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimProperty"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimProperty
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimNone = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimX = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimY = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimWidth = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimHeight = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimOpacity = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimRotation = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimColor = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimVisibility = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontBold = 100,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontColor = 101,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontEmboss = 102,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontItalic = 103,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontName = 104,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontShadow = 105,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontSize = 106,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontSubscript = 107,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontSuperscript = 108,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontUnderline = 109,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextFontStrikeThrough = 110,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletCharacter = 111,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletFontName = 112,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletNumber = 113,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletColor = 114,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletRelativeSize = 115,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletStyle = 116,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimTextBulletType = 117,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapePictureContrast = 1000,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapePictureBrightness = 1001,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1002</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapePictureGamma = 1002,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1003</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapePictureGrayscale = 1003,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1004</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeFillOn = 1004,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1005</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeFillColor = 1005,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1006</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeFillOpacity = 1006,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeFillBackColor = 1007,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1008</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeLineOn = 1008,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1009</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeLineColor = 1009,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1010</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowOn = 1010,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1011</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowType = 1011,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1012</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowColor = 1012,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1013</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowOpacity = 1013,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1014</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowOffsetX = 1014,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1015</remarks>
		 [SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		 msoAnimShapeShadowOffsetY = 1015
	}
}