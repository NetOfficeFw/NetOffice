using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11
	 /// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsEnum)]
	public enum PjCacheJobState
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateInvalid = -1,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateUnknown = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateReadyForProcessing = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateSendIncomplete = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateProcessing = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateSuccess = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateFailed = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateFailedNotBlocking = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateSkipped = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateCorrelationBlocked = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateCancelled = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateOnHold = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateSleeping = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateReadyForLaunch = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11)]
		 pjCacheJobStateLastState = 13
	}
}