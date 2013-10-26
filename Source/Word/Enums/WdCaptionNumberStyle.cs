using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839097.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCaptionNumberStyle
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleArabic = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleUppercaseRoman = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleLowercaseRoman = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleUppercaseLetter = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleLowercaseLetter = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleArabicFullWidth = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleKanji = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleKanjiDigit = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleKanjiTraditional = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleNumberInCircle = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleGanada = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleChosung = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleZodiac1 = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleZodiac2 = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleHanjaRead = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleHanjaReadDigit = 42,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleTradChinNum2 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleTradChinNum3 = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleSimpChinNum2 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleSimpChinNum3 = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleHebrewLetter1 = 45,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleArabicLetter1 = 46,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleHebrewLetter2 = 47,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdCaptionNumberStyleArabicLetter2 = 48,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleHindiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleHindiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleHindiArabic = 51,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleHindiCardinalText = 52,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleThaiLetter = 53,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleThaiArabic = 54,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleThaiCardinalText = 55,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdCaptionNumberStyleVietCardinalText = 56
	}
}