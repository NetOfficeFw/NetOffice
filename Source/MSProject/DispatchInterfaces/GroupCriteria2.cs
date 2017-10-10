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
	/// DispatchInterface GroupCriteria2 
	/// SupportByVersion MSProject, 11,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920608(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class GroupCriteria2 : COMObject, IEnumerableProvider<NetOffice.MSProjectApi.GroupCriterion2>
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public GroupCriteria2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public GroupCriteria2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupCriteria2(string progId) : base(progId)
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
		public NetOffice.MSProjectApi.GroupCriterion2 this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Item", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Group2 Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Group2>(this, "Parent", NetOffice.MSProjectApi.Group2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt, object groupInterval)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt, groupInterval });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending, fontName);
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending, fontName, fontSize);
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern });
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
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn });
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
		public NetOffice.MSProjectApi.GroupCriterion2 Add(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "Add", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt, object groupInterval)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt, groupInterval });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="fieldName">string fieldName</param>
		/// <param name="ascending">optional bool Ascending = true</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending, fontName);
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, fieldName, ascending, fontName, fontSize);
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn });
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
		public NetOffice.MSProjectApi.GroupCriterion2 AddEx(string fieldName, object ascending, object fontName, object fontSize, object fontBold, object fontItalic, object fontUnderLine, object fontColor, object cellColor, object pattern, object groupOn, object startAt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.GroupCriterion2>(this, "AddEx", NetOffice.MSProjectApi.GroupCriterion2.LateBindingApiWrapperType, new object[]{ fieldName, ascending, fontName, fontSize, fontBold, fontItalic, fontUnderLine, fontColor, cellColor, pattern, groupOn, startAt });
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
        public IEnumerator<NetOffice.MSProjectApi.GroupCriterion2> GetEnumerator()
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