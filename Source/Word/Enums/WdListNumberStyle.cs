using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837344.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdListNumberStyle
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleArabic = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleUppercaseRoman = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleLowercaseRoman = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleUppercaseLetter = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleLowercaseLetter = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleOrdinal = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleCardinalText = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleOrdinalText = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleKanji = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleKanjiDigit = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleAiueoHalfWidth = 12,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleIrohaHalfWidth = 13,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleArabicFullWidth = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleKanjiTraditional = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleKanjiTraditional2 = 17,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleNumberInCircle = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleAiueo = 20,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleIroha = 21,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleArabicLZ = 22,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleBullet = 23,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleGanada = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleChosung = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleGBNum1 = 26,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleGBNum2 = 27,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleGBNum3 = 28,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleGBNum4 = 29,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleZodiac1 = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleZodiac2 = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleZodiac3 = 32,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleTradChinNum1 = 33,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleTradChinNum2 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleTradChinNum3 = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleTradChinNum4 = 36,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleSimpChinNum1 = 37,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleSimpChinNum2 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleSimpChinNum3 = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleSimpChinNum4 = 40,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHanjaRead = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHanjaReadDigit = 42,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHangul = 43,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHanja = 44,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHebrew1 = 45,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleArabic1 = 46,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleHebrew2 = 47,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleArabic2 = 48,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>253</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleLegal = 253,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>254</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleLegalLZ = 254,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>255</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdListNumberStyleNone = 255,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleHindiLetter1 = 49,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleHindiLetter2 = 50,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleHindiArabic = 51,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleHindiCardinalText = 52,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleThaiLetter = 53,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleThaiArabic = 54,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleThaiCardinalText = 55,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleVietCardinalText = 56,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleLowercaseRussian = 58,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStyleUppercaseRussian = 59,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>249</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdListNumberStylePictureBullet = 249,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleLowercaseGreek = 60,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleUppercaseGreek = 61,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleArabicLZ2 = 62,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleArabicLZ3 = 63,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleArabicLZ4 = 64,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleLowercaseTurkish = 65,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleUppercaseTurkish = 66,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleLowercaseBulgarian = 67,

		 /// <summary>
		 /// SupportByVersion Word 14, 15, 16
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("Word", 14,15,16)]
		 wdListNumberStyleUppercaseBulgarian = 68
	}
}