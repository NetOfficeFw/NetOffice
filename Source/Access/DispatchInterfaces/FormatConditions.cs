using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface FormatConditions 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845092.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Access", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("E27A992E-A330-11D0-81DD-00C04FC2F51B")]
	public interface FormatConditions : ICOMObject, IEnumerableProvider<NetOffice.AccessApi._FormatCondition>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192090.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198351.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.AccessApi._FormatCondition this[object index] { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192647.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194814.aspx </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator operator = 0</param>
		/// <param name="expression1">optional object expression1</param>
		/// <param name="expression2">optional object expression2</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator, object expression1, object expression2);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194814.aspx </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194814.aspx </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194814.aspx </remarks>
		/// <param name="type">NetOffice.AccessApi.Enums.AcFormatConditionType type</param>
		/// <param name="_operator">optional NetOffice.AccessApi.Enums.AcFormatConditionOperator Operator = 0</param>
		/// <param name="expression1">optional object expression1</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi._FormatCondition Add(NetOffice.AccessApi.Enums.AcFormatConditionType type, object _operator, object expression1);

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836089.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

        #endregion

        #region IEnumerable<NetOffice.AccessApi._FormatCondition>

        /// <summary>
        /// SupportByVersion Access, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.AccessApi._FormatCondition> GetEnumerator();

        #endregion
    }
}
