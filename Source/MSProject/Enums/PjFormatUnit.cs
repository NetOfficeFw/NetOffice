using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863211(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjFormatUnit
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMinutes = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjHours = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjDays = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjWeeks = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonths = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedMinutes = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedHours = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedDays = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedWeeks = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedMonths = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMinutesEstimated = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjHoursEstimated = 37,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjDaysEstimated = 39,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjWeeksEstimated = 41,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjMonthsEstimated = 43,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedMinutesEstimated = 36,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedHoursEstimated = 38,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedDaysEstimated = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedWeeksEstimated = 42,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjElapsedMonthsEstimated = 44
	}
}