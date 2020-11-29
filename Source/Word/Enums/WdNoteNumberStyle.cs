﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdNoteNumberStyle"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdNoteNumberStyle
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleArabic = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleUppercaseRoman = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleLowercaseRoman = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleUppercaseLetter = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleLowercaseLetter = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleSymbol = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleArabicFullWidth = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleKanji = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleKanjiDigit = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleKanjiTraditional = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleNumberInCircle = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleHanjaRead = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleHanjaReadDigit = 42,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleTradChinNum1 = 33,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleTradChinNum2 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleSimpChinNum1 = 37,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleSimpChinNum2 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleHebrewLetter1 = 45,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleArabicLetter1 = 46,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleHebrewLetter2 = 47,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdNoteNumberStyleArabicLetter2 = 48,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleHindiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleHindiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleHindiArabic = 51,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleHindiCardinalText = 52,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleThaiLetter = 53,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleThaiArabic = 54,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleThaiCardinalText = 55,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdNoteNumberStyleVietCardinalText = 56
	}
}