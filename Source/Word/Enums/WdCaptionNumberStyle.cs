﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdCaptionNumberStyle"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdCaptionNumberStyle
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleArabic = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleUppercaseRoman = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleLowercaseRoman = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleUppercaseLetter = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleLowercaseLetter = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleArabicFullWidth = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleKanji = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleKanjiDigit = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleKanjiTraditional = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleNumberInCircle = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleGanada = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleChosung = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleZodiac1 = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleZodiac2 = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleHanjaRead = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleHanjaReadDigit = 42,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleTradChinNum2 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleTradChinNum3 = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleSimpChinNum2 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleSimpChinNum3 = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleHebrewLetter1 = 45,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleArabicLetter1 = 46,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleHebrewLetter2 = 47,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdCaptionNumberStyleArabicLetter2 = 48,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleHindiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleHindiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleHindiArabic = 51,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleHindiCardinalText = 52,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleThaiLetter = 53,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleThaiArabic = 54,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleThaiCardinalText = 55,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdCaptionNumberStyleVietCardinalText = 56
	}
}