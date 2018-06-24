using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface GroupCriteria 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920605(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11, 12, 14), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("11589052-69F0-11D2-B889-00C04FB90729")]
	public interface GroupCriteria : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.GroupCriterion>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.GroupCriterion this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Group Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		/// <param name="cellColor">optional NetOffice.MSProjectApi.Enums.PjColor CellColor = 16</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		/// <param name="startAt">optional object StartAt = 0</param>
		/// <param name="groupInterval">optional object GroupInterval = 1</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt, object groupInterval);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		/// <param name="cellColor">optional NetOffice.MSProjectApi.Enums.PjColor CellColor = 16</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		/// <param name="cellColor">optional NetOffice.MSProjectApi.Enums.PjColor CellColor = 16</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		/// <param name="cellColor">optional NetOffice.MSProjectApi.Enums.PjColor CellColor = 16</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional NetOffice.MSProjectApi.Enums.PjColor FontColor = 0</param>
		/// <param name="cellColor">optional NetOffice.MSProjectApi.Enums.PjColor CellColor = 16</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		/// <param name="startAt">optional object StartAt = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.GroupCriterion Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.GroupCriterion>

        /// <summary>
        /// SupportByVersion MSProject, 11,12,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 12, 14)]
        new IEnumerator<NetOffice.MSProjectApi.GroupCriterion> GetEnumerator();

        #endregion
    }
}
