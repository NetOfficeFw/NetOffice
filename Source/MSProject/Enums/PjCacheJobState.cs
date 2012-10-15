using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjCacheJobState
	{
		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateInvalid = -1,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateUnknown = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateReadyForProcessing = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateSendIncomplete = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateProcessing = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateSuccess = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateFailed = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateFailedNotBlocking = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateSkipped = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateCorrelationBlocked = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateCancelled = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateOnHold = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateSleeping = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateReadyForLaunch = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjCacheJobStateLastState = 13
	}
}