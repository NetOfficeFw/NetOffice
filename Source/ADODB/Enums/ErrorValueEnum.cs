using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum ErrorValueEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3001</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrInvalidArgument = 3001,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3021</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrNoCurrentRecord = 3021,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3219</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrIllegalOperation = 3219,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3246</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrInTransaction = 3246,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3251</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrFeatureNotAvailable = 3251,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3265</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrItemNotFound = 3265,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3367</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrObjectInCollection = 3367,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3420</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrObjectNotSet = 3420,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3421</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrDataConversion = 3421,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3704</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrObjectClosed = 3704,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3705</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrObjectOpen = 3705,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3706</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrProviderNotFound = 3706,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3707</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrBoundToCommand = 3707,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3708</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrInvalidParamInfo = 3708,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3709</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrInvalidConnection = 3709,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3710</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrNotReentrant = 3710,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3711</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrStillExecuting = 3711,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3712</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrOperationCancelled = 3712,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3713</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrStillConnecting = 3713,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3715</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrNotExecuting = 3715,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3716</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adErrUnsafeOperation = 3716,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3000</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrProviderFailed = 3000,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3002</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrOpeningFile = 3002,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3003</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrReadFile = 3003,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3004</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrWriteFile = 3004,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3220</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCantChangeProvider = 3220,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3714</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrInvalidTransaction = 3714,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3717</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adwrnSecurityDialog = 3717,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3718</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adwrnSecurityDialogHeader = 3718,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3719</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrIntegrityViolation = 3719,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3720</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPermissionDenied = 3720,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3721</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrDataOverflow = 3721,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3722</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrSchemaViolation = 3722,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3723</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrSignMismatch = 3723,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3724</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCantConvertvalue = 3724,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3725</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCantCreate = 3725,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3726</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrColumnNotOnThisRow = 3726,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3727</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrURLDoesNotExist = 3727,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3728</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrTreePermissionDenied = 3728,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3729</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrInvalidURL = 3729,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3730</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrResourceLocked = 3730,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3731</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrResourceExists = 3731,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3732</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCannotComplete = 3732,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3733</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrVolumeNotFound = 3733,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3734</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrOutOfSpace = 3734,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3735</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrResourceOutOfScope = 3735,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3736</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrUnavailable = 3736,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3737</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrURLNamedRowDoesNotExist = 3737,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3738</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrDelResOutOfScope = 3738,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3739</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropInvalidColumn = 3739,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3740</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropInvalidOption = 3740,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3741</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropInvalidValue = 3741,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3742</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropConflicting = 3742,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3743</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropNotAllSettable = 3743,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3744</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropNotSet = 3744,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3745</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropNotSettable = 3745,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3746</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrPropNotSupported = 3746,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3747</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCatalogNotSet = 3747,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3748</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrCantChangeConnection = 3748,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3749</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrFieldsUpdateFailed = 3749,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3750</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrDenyNotSupported = 3750,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3751</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adErrDenyTypeNotSupported = 3751
	}
}