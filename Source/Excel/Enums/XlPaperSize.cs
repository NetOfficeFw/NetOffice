﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPaperSize"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlPaperSize
	{
		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaper10x14 = 16,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaper11x17 = 17,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperA3 = 8,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperA4 = 9,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperA4Small = 10,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperA5 = 11,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperB4 = 12,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperB5 = 13,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperCsheet = 24,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperDsheet = 25,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelope10 = 20,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelope11 = 21,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelope12 = 22,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelope14 = 23,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelope9 = 19,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeB4 = 33,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeB5 = 34,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeB6 = 35,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeC3 = 29,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeC4 = 30,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeC5 = 28,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeC6 = 31,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeC65 = 32,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeDL = 27,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeItaly = 36,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopeMonarch = 37,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEnvelopePersonal = 38,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperEsheet = 26,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperExecutive = 7,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperFanfoldLegalGerman = 41,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperFanfoldStdGerman = 40,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperFanfoldUS = 39,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperFolio = 14,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperLedger = 4,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperLegal = 5,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperLetter = 1,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperLetterSmall = 2,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperNote = 18,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperQuarto = 15,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperStatement = 6,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperTabloid = 3,

		 /// <summary>
		 /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		 xlPaperUser = 256
	}
}