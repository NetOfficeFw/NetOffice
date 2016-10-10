using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864118.aspx </remarks>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoChartElementType
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartTitleNone = 0,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartTitleCenteredOverlay = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartTitleAboveChart = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendNone = 100,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendRight = 101,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendTop = 102,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendLeft = 103,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendBottom = 104,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendRightOverlay = 105,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLegendLeftOverlay = 106,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelNone = 200,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelShow = 201,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelCenter = 202,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelInsideEnd = 203,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelInsideBase = 204,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelOutSideEnd = 205,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelLeft = 206,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>207</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelRight = 207,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelTop = 208,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelBottom = 209,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>210</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataLabelBestFit = 210,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleNone = 300,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>301</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleAdjacentToAxis = 301,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>302</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleBelowAxis = 302,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>303</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleRotated = 303,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>304</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleVertical = 304,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>305</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisTitleHorizontal = 305,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>306</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleNone = 306,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>306</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleAdjacentToAxis = 306,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>308</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleBelowAxis = 308,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>309</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleRotated = 309,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>310</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleVertical = 310,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>311</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisTitleHorizontal = 311,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>312</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleNone = 312,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>313</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleAdjacentToAxis = 313,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>314</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleBelowAxis = 314,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>315</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleRotated = 315,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>316</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleVertical = 316,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>317</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisTitleHorizontal = 317,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>318</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleNone = 318,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>319</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleAdjacentToAxis = 319,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>320</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleBelowAxis = 320,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>321</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleRotated = 321,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>322</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleVertical = 322,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>323</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisTitleHorizontal = 323,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>324</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisTitleNone = 324,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>325</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisTitleRotated = 325,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>326</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisTitleVertical = 326,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>327</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisTitleHorizontal = 327,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>328</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueGridLinesNone = 328,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>329</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueGridLinesMinor = 329,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>330</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueGridLinesMajor = 330,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>331</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueGridLinesMinorMajor = 331,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>332</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryGridLinesNone = 332,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>333</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryGridLinesMinor = 333,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>334</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryGridLinesMajor = 334,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>335</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryGridLinesMinorMajor = 335,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>336</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueGridLinesNone = 336,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>337</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueGridLinesMinor = 337,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>338</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueGridLinesMajor = 338,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>339</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueGridLinesMinorMajor = 339,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>340</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryGridLinesNone = 340,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>341</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryGridLinesMinor = 341,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>342</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryGridLinesMajor = 342,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>343</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryGridLinesMinorMajor = 343,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>344</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisGridLinesNone = 344,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>345</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisGridLinesMinor = 345,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>346</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisGridLinesMajor = 346,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>347</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisGridLinesMinorMajor = 347,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>348</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisNone = 348,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>349</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisShow = 349,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>350</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisWithoutLabels = 350,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>351</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisReverse = 351,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>352</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisNone = 352,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisShow = 353,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>354</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisThousands = 354,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>355</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisMillions = 355,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>356</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisBillions = 356,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>357</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryValueAxisLogScale = 357,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>358</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisNone = 358,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>359</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisShow = 359,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>360</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisWithoutLabels = 360,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>361</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisReverse = 361,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>362</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisNone = 362,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>363</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisShow = 363,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>364</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisThousands = 364,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>365</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisMillions = 365,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>366</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisBillions = 366,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>367</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryValueAxisLogScale = 367,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>368</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisNone = 368,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>369</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisShow = 369,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>370</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisWithoutLabeling = 370,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>371</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSeriesAxisReverse = 371,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>372</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisThousands = 372,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>373</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisMillions = 373,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>374</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisBillions = 374,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>375</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPrimaryCategoryAxisLogScale = 375,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>376</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisThousands = 376,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>377</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisMillions = 377,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>378</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisBillions = 378,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>379</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementSecondaryCategoryAxisLogScale = 379,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataTableNone = 500,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataTableShow = 501,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>502</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementDataTableWithLegendKeys = 502,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>600</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementTrendlineNone = 600,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>601</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementTrendlineAddLinear = 601,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>602</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementTrendlineAddExponential = 602,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>603</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementTrendlineAddLinearForecast = 603,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>604</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementTrendlineAddTwoPeriodMovingAverage = 604,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>700</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementErrorBarNone = 700,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementErrorBarStandardError = 701,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementErrorBarPercentage = 702,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>703</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementErrorBarStandardDeviation = 703,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>800</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLineNone = 800,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>801</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLineDropLine = 801,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>802</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLineHiLoLine = 802,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>803</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLineSeriesLine = 803,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>804</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementLineDropHiLoLine = 804,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>900</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementUpDownBarsNone = 900,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>901</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementUpDownBarsShow = 901,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPlotAreaNone = 1000,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementPlotAreaShow = 1001,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1100</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartWallNone = 1100,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1101</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartWallShow = 1101,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartFloorNone = 1200,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1201</remarks>
		 [SupportByVersionAttribute("Office", 12,14,15,16)]
		 msoElementChartFloorShow = 1201,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoElementDataLabelCallout = 211
	}
}