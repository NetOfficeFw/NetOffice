﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlChartType"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlChartType
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlColumnClustered = 51,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlColumnStacked = 52,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlColumnStacked100 = 53,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DColumnClustered = 54,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DColumnStacked = 55,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DColumnStacked100 = 56,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBarClustered = 57,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBarStacked = 58,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBarStacked100 = 59,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DBarClustered = 60,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DBarStacked = 61,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DBarStacked100 = 62,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLineStacked = 63,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLineStacked100 = 64,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLineMarkers = 65,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLineMarkersStacked = 66,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLineMarkersStacked100 = 67,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPieOfPie = 68,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPieExploded = 69,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DPieExploded = 70,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBarOfPie = 71,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlXYScatterSmooth = 72,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlXYScatterSmoothNoMarkers = 73,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlXYScatterLines = 74,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlXYScatterLinesNoMarkers = 75,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlAreaStacked = 76,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlAreaStacked100 = 77,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DAreaStacked = 78,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DAreaStacked100 = 79,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlDoughnutExploded = 80,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlRadarMarkers = 81,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlRadarFilled = 82,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSurface = 83,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSurfaceWireframe = 84,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSurfaceTopView = 85,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlSurfaceTopViewWireframe = 86,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBubble = 15,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlBubble3DEffect = 87,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStockHLC = 88,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStockOHLC = 89,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStockVHLC = 90,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlStockVOHLC = 91,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderColClustered = 92,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderColStacked = 93,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderColStacked100 = 94,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderBarClustered = 95,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderBarStacked = 96,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderBarStacked100 = 97,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlCylinderCol = 98,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeColClustered = 99,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeColStacked = 100,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeColStacked100 = 101,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeBarClustered = 102,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeBarStacked = 103,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeBarStacked100 = 104,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlConeCol = 105,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidColClustered = 106,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidColStacked = 107,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidColStacked100 = 108,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidBarClustered = 109,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidBarStacked = 110,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidBarStacked100 = 111,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPyramidCol = 112,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4100</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DColumn = -4100,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlLine = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4101</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DLine = -4101,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4102</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DPie = -4102,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPie = 5,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4169</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlXYScatter = -4169,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4098</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xl3DArea = -4098,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlArea = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4120</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlDoughnut = -4120,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4151</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlRadar = -4151
	}
}