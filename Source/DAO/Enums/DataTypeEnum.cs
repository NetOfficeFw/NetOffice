using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 3.6, 12.0
	 /// </summary>
	[SupportByVersionAttribute("DAO", 3.6,12.0)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum DataTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbBoolean = 1,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbByte = 2,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbInteger = 3,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbLong = 4,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbCurrency = 5,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbSingle = 6,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbDouble = 7,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbDate = 8,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbBinary = 9,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbText = 10,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbLongBinary = 11,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbMemo = 12,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbGUID = 15,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbBigInt = 16,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbVarBinary = 17,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbChar = 18,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbNumeric = 19,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbDecimal = 20,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbFloat = 21,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbTime = 22,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("DAO", 3.6,12.0)]
		 dbTimeStamp = 23,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbAttachment = 101,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexByte = 102,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexInteger = 103,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexLong = 104,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexSingle = 105,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexDouble = 106,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexGUID = 107,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexDecimal = 108,

		 /// <summary>
		 /// SupportByVersion DAO 12.0
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersionAttribute("DAO", 12.0)]
		 dbComplexText = 109
	}
}