using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface GroupCriteria2 
	/// SupportByVersion MSProject, 11,14
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920608(v=office.14).aspx </remarks>
	public class GroupCriteria2 : COMObject, NetOffice.MSProjectApi.GroupCriteria2
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSProjectApi.GroupCriteria2);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(GroupCriteria2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public GroupCriteria2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSProject", 11,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Item", typeof(NetOffice.MSProjectApi.GroupCriterion2), index);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.Group2 Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Group2>(this, "Parent", typeof(NetOffice.MSProjectApi.Group2));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", typeof(NetOffice.MSProjectApi.Application));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt, object groupInterval)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt, groupInterval });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending, fontName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending, fontName, fontSize);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern });		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
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
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		/// <param name="cellColor">optional Int32 CellColor = -16777216</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		/// <param name="startAt">optional object StartAt = 0</param>
		/// <param name="groupInterval">optional object GroupInterval = 1</param>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt, object groupInterval)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt, groupInterval });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending, fontName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), fieldName, ascending, fontName, fontSize);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		/// <param name="cellColor">optional Int32 CellColor = -16777216</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		/// <param name="cellColor">optional Int32 CellColor = -16777216</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		/// <param name="cellColor">optional Int32 CellColor = -16777216</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional Int32 FontSize = 8</param>
		/// <param name="fontBold">optional bool FontBold = true</param>
		/// <param name="fontItalic">optional bool FontItalic = false</param>
		/// <param name="fontUnderLine">optional bool FontUnderLine = false</param>
		/// <param name="fontColor">optional Int32 FontColor = -16777216</param>
		/// <param name="cellColor">optional Int32 CellColor = -16777216</param>
		/// <param name="pattern">optional NetOffice.MSProjectApi.Enums.PjBackgroundPattern Pattern = -1</param>
		/// <param name="groupOn">optional NetOffice.MSProjectApi.Enums.PjGroupOn GroupOn = 0</param>
		/// <param name="startAt">optional object StartAt = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", typeof(NetOffice.MSProjectApi.GroupCriterion2), new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.GroupCriterion2>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.GroupCriterion2>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.GroupCriterion2>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.GroupCriterion2>

        /// <summary>
        /// SupportByVersion MSProject, 11,14
        /// </summary>
        [SupportByVersion("MSProject", 11, 14)]
        public virtual IEnumerator<NetOffice.MSProjectApi.GroupCriterion2> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.GroupCriterion2 item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSProject, 11,14
        /// </summary>
        [SupportByVersion("MSProject", 11,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

