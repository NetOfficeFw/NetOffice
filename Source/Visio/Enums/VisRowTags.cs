using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRowTags
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagDefault = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagBase = 130,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagRowVoid = 180,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagInvalid = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagComponent = 137,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagMoveTo = 138,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagLineTo = 139,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagArcTo = 140,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagInfiniteLine = 141,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>143</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagEllipse = 143,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagEllipticalArcTo = 144,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagSplineBeg = 165,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagSplineSpan = 166,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>193</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagPolylineTo = 193,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>195</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagNURBSTo = 195,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagTab0 = 136,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagTab2 = 150,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagTab10 = 151,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagTab60 = 181,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCnnctPt = 153,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>185</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCnnctNamed = 185,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>186</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCnnctPtABCD = 186,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>187</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCnnctNamedABCD = 187,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCtlPt = 162,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visTagCtlPtTip = 170,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>238</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visTagRelMoveTo = 238,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>239</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visTagRelLineTo = 239,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visTagRelEllipticalArcTo = 240,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>236</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visTagRelCubBezTo = 236,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>237</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visTagRelQuadBezTo = 237
	}
}