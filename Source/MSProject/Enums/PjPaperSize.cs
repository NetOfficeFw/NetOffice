using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867028(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjPaperSize
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaper10x14 = 16,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaper11x17 = 17,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperA3 = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperA4 = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperA4Small = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperA5 = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperB4 = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperB5 = 13,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperCsheet = 24,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperDsheet = 25,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelop10 = 20,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelope11 = 21,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelope12 = 22,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelope14 = 23,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelope9 = 19,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeB4 = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeB5 = 34,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopB6 = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeC3 = 29,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeC4 = 30,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeC5 = 28,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeC6 = 31,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeC65 = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeDL = 27,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeItaly = 36,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopeMonarch = 37,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEnvelopePersonal = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperEsheet = 26,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperExecutive = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperFanfoldLegalGerman = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperFanfoldStdGerman = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperFanfoldUS = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperFolio = 14,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperLedger = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperLegal = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperLetter = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperLetterSmall = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperNote = 18,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperQuarto = 15,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperStatement = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperTabloid = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>255</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperUser = 255,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjPaperDefault = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperISOB4 = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapanesePostcard = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaper9X11 = 44,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaper10X11 = 45,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaper15x11 = 46,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperEnvelopeInvite = 47,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLetterExtra = 50,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLegalExtra = 51,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperTabloidExtra = 52,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA4Extra = 53,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLetterTransverse = 54,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA4Transverse = 55,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLetterExtraTransverse = 56,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperSuperA = 57,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperSuperB = 58,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLetterPlus = 59,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA4Plus = 60,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA5Transverse = 61,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB5Transverse = 62,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA3Extra = 63,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA5Extra = 64,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB5Extra = 65,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA2 = 66,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA3Transverse = 67,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA3ExtraTransverse = 68,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseDoublePostcard = 69,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA6 = 70,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeKaku2 = 71,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeKaku3 = 72,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeChou3 = 73,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeChou4 = 74,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperLetterRotated = 75,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA3Rotated = 76,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA4Rotated = 77,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA5Rotated = 78,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB4JISRotated = 79,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB5JISRotated = 80,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapanesePostcardRotated = 81,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperDoubleJapanesePostcardRotated = 82,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperA6Rotated = 83,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeKaku2Rotated = 84,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeKaku3Rotated = 85,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeChou3Rotated = 86,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeChou4Rotated = 87,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB6JIS = 88,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperB6JISRotated = 89,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaper12x11 = 90,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeYou4 = 91,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperJapaneseEnvelopeYou4Rotated = 92,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRC16K = 93,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRC32K = 94,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRC32KBig = 95,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope1 = 96,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope2 = 97,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope3 = 98,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope4 = 99,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope5 = 100,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope6 = 101,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope7 = 102,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope8 = 103,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope9 = 104,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope10 = 105,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope16K = 106,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope32K = 107,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope32KBigRotated = 108,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope1Rotated = 109,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope2Rotated = 110,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope3Rotated = 111,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope4Rotated = 112,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope5Rotated = 113,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope6Rotated = 114,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope7Rotated = 115,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope8Rotated = 116,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope9Rotated = 117,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjPaperPRCEnvelope10Rotated = 118
	}
}