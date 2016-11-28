using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum FieldStatusEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldOK = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldCantConvertValue = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldIsNull = 3,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldTruncated = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldSignMismatch = 5,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldDataOverflow = 6,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldCantCreate = 7,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldUnavailable = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPermissionDenied = 9,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldIntegrityViolation = 10,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldSchemaViolation = 11,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldBadStatus = 12,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldDefault = 13,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldIgnore = 15,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldDoesNotExist = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldInvalidURL = 17,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldResourceLocked = 18,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldResourceExists = 19,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldCannotComplete = 20,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldVolumeNotFound = 21,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldOutOfSpace = 22,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldCannotDeleteSource = 23,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldReadOnly = 24,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldResourceOutOfScope = 25,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldAlreadyExists = 26,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>65536</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPendingInsert = 65536,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>131072</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPendingDelete = 131072,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>262144</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPendingChange = 262144,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>524288</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPendingUnknown = 524288,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>1048576</remarks>
		 [SupportByVersionAttribute("ADODB", 2.5)]
		 adFieldPendingUnknownDelete = 1048576
	}
}