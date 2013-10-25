using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862770.aspx </remarks>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoAutoShapeType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRectangle = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeParallelogram = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeTrapezoid = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDiamond = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRoundedRectangle = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeOctagon = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeIsoscelesTriangle = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRightTriangle = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeOval = 9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeHexagon = 10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCross = 11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRegularPentagon = 12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCan = 13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCube = 14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeBevel = 15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFoldedCorner = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeSmileyFace = 17,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDonut = 18,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeNoSymbol = 19,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeBlockArc = 20,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeHeart = 21,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLightningBolt = 22,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeSun = 23,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeMoon = 24,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeArc = 25,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDoubleBracket = 26,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDoubleBrace = 27,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapePlaque = 28,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftBracket = 29,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRightBracket = 30,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftBrace = 31,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRightBrace = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRightArrow = 33,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftArrow = 34,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUpArrow = 35,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDownArrow = 36,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftRightArrow = 37,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUpDownArrow = 38,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeQuadArrow = 39,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftRightUpArrow = 40,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeBentArrow = 41,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUTurnArrow = 42,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftUpArrow = 43,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeBentUpArrow = 44,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedRightArrow = 45,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedLeftArrow = 46,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedUpArrow = 47,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedDownArrow = 48,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeStripedRightArrow = 49,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeNotchedRightArrow = 50,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapePentagon = 51,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeChevron = 52,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRightArrowCallout = 53,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftArrowCallout = 54,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUpArrowCallout = 55,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDownArrowCallout = 56,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLeftRightArrowCallout = 57,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUpDownArrowCallout = 58,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeQuadArrowCallout = 59,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCircularArrow = 60,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartProcess = 61,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartAlternateProcess = 62,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartDecision = 63,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartData = 64,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartPredefinedProcess = 65,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartInternalStorage = 66,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartDocument = 67,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartMultidocument = 68,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartTerminator = 69,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartPreparation = 70,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartManualInput = 71,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartManualOperation = 72,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartConnector = 73,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartOffpageConnector = 74,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartCard = 75,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartPunchedTape = 76,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartSummingJunction = 77,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartOr = 78,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartCollate = 79,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartSort = 80,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartExtract = 81,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartMerge = 82,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartStoredData = 83,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartDelay = 84,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartSequentialAccessStorage = 85,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartMagneticDisk = 86,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartDirectAccessStorage = 87,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeFlowchartDisplay = 88,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeExplosion1 = 89,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeExplosion2 = 90,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape4pointStar = 91,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape5pointStar = 92,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape8pointStar = 93,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape16pointStar = 94,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape24pointStar = 95,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShape32pointStar = 96,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeUpRibbon = 97,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDownRibbon = 98,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedUpRibbon = 99,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCurvedDownRibbon = 100,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeVerticalScroll = 101,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeHorizontalScroll = 102,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeWave = 103,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeDoubleWave = 104,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRectangularCallout = 105,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeRoundedRectangularCallout = 106,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeOvalCallout = 107,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeCloudCallout = 108,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout1 = 109,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout2 = 110,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout3 = 111,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout4 = 112,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout1AccentBar = 113,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout2AccentBar = 114,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout3AccentBar = 115,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout4AccentBar = 116,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout1NoBorder = 117,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout2NoBorder = 118,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout3NoBorder = 119,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout4NoBorder = 120,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>121</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout1BorderandAccentBar = 121,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout2BorderandAccentBar = 122,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout3BorderandAccentBar = 123,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeLineCallout4BorderandAccentBar = 124,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonCustom = 125,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonHome = 126,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonHelp = 127,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonInformation = 128,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonBackorPrevious = 129,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonForwardorNext = 130,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonBeginning = 131,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonEnd = 132,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonReturn = 133,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonDocument = 134,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonSound = 135,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeActionButtonMovie = 136,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeBalloon = 137,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoShapeNotPrimitive = 138,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeFlowchartOfflineStorage = 139,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeLeftRightRibbon = 140,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeDiagonalStripe = 141,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapePie = 142,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>143</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeNonIsoscelesTrapezoid = 143,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeDecagon = 144,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>145</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeHeptagon = 145,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>146</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeDodecagon = 146,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>147</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShape6pointStar = 147,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>148</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShape7pointStar = 148,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>149</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShape10pointStar = 149,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShape12pointStar = 150,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeRound1Rectangle = 151,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeRound2SameRectangle = 152,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeRound2DiagRectangle = 153,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>154</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSnipRoundRectangle = 154,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>155</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSnip1Rectangle = 155,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>156</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSnip2SameRectangle = 156,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>157</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSnip2DiagRectangle = 157,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>158</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeFrame = 158,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeHalfFrame = 159,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeTear = 160,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>161</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeChord = 161,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeCorner = 162,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathPlus = 163,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathMinus = 164,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathMultiply = 165,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathDivide = 166,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathEqual = 167,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeMathNotEqual = 168,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeCornerTabs = 169,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSquareTabs = 170,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapePlaqueTabs = 171,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeGear6 = 172,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeGear9 = 173,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>174</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeFunnel = 174,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapePieWedge = 175,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeLeftCircularArrow = 176,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeLeftRightCircularArrow = 177,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeSwooshArrow = 178,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeCloud = 179,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeChartX = 180,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeChartStar = 181,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeChartPlus = 182,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15
		 /// </summary>
		 /// <remarks>183</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15)]
		 msoShapeLineInverse = 183
	}
}