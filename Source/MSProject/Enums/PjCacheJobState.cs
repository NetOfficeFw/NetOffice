using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 16
	 /// </summary>
	[SupportByVersion("MSProject", 11,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PjCacheJobState
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateInvalid = -1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateUnknown = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateReadyForProcessing = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateSendIncomplete = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateProcessing = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateSuccess = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateFailed = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateFailedNotBlocking = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateSkipped = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateCorrelationBlocked = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateCancelled = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateOnHold = 10,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateSleeping = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateReadyForLaunch = 12,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("MSProject", 11,16)]
		 pjCacheJobStateLastState = 13
	}
}