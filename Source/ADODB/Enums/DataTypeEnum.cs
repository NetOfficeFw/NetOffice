using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum DataTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adEmpty = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adTinyInt = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSmallInt = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adInteger = 3,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adBigInt = 20,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adUnsignedTinyInt = 17,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adUnsignedSmallInt = 18,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adUnsignedInt = 19,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adUnsignedBigInt = 21,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSingle = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDouble = 5,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adCurrency = 6,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDecimal = 14,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adNumeric = 131,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adBoolean = 11,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adError = 10,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adUserDefined = 132,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adVariant = 12,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adIDispatch = 9,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adIUnknown = 13,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adGUID = 72,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDate = 7,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDBDate = 133,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDBTime = 134,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adDBTimeStamp = 135,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adBSTR = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adChar = 129,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adVarChar = 200,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adLongVarChar = 201,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adWChar = 130,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adVarWChar = 202,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adLongVarWChar = 203,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adBinary = 128,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adVarBinary = 204,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adLongVarBinary = 205,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adChapter = 136,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adFileTime = 64,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersion("ADODB", 2.1)]
		 adDBFileTime = 137,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adPropVariant = 138,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adVarNumeric = 139,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adArray = 8192
	}
}