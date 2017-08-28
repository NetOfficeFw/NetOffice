using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotFieldSet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotFieldSet : COMObject
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
                    _type = typeof(PivotFieldSet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotFieldSet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotFieldSet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFieldSet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFields>(this, "Fields", NetOffice.OWC10Api.PivotFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotMembers Members
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotMembers>(this, "Members", NetOffice.OWC10Api.PivotMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldSetOrientationEnum Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldSetOrientationEnum>(this, "Orientation");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldSetTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldSetTypeEnum>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField BoundField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotField>(this, "BoundField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string UniqueName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.PivotMember DefaultMember
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.PivotMember>(this, "DefaultMember");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.PivotMember Member
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.PivotMember>(this, "Member");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.PivotMember AllMember
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.PivotMember>(this, "AllMember");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum CompareOrderedMembersBy
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum>(this, "CompareOrderedMembersBy");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CompareOrderedMembersBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotView View
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "View", NetOffice.OWC10Api.PivotView.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotFilterUpdate CreateFilterUpdate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFilterUpdate>(this, "CreateFilterUpdate", NetOffice.OWC10Api.PivotFilterUpdate.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowMultiFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowMultiFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowMultiFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string FilterCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FilterCaption");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldSetAllIncludeExcludeEnum AllIncludeExclude
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldSetAllIncludeExcludeEnum>(this, "AllIncludeExclude");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AllIncludeExclude", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum CompareMemberCaptionsBy
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum>(this, "CompareMemberCaptionsBy");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CompareMemberCaptionsBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayInFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayInFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayInFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AlwaysIncludeInCube
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlwaysIncludeInCube");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlwaysIncludeInCube", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotMember get_FindMember(object nameOrPath, object format)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotMember>(this, "FindMember", NetOffice.OWC10Api.PivotMember.LateBindingApiWrapperType, nameOrPath, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindMember
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindMember")]
		public NetOffice.OWC10Api.PivotMember FindMember(object nameOrPath, object format)
		{
			return get_FindMember(nameOrPath, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotMember get_FindMember(object nameOrPath)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotMember>(this, "FindMember", NetOffice.OWC10Api.PivotMember.LateBindingApiWrapperType, nameOrPath);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindMember
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindMember")]
		public NetOffice.OWC10Api.PivotMember FindMember(object nameOrPath)
		{
			return get_FindMember(nameOrPath);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="dataField">string dataField</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField AddCalculatedField(string name, string caption, string dataField, string expression)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotField>(this, "AddCalculatedField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType, name, caption, dataField, expression);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="caption">optional string Caption = </param>
		/// <param name="before">optional object Before = 0</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField AddCustomGroupField(object name, object caption, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotField>(this, "AddCustomGroupField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType, name, caption, before);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField AddCustomGroupField()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotField>(this, "AddCustomGroupField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField AddCustomGroupField(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotField>(this, "AddCustomGroupField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="caption">optional string Caption = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField AddCustomGroupField(object name, object caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotField>(this, "AddCustomGroupField", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType, name, caption);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">object field</param>
		[SupportByVersion("OWC10", 1)]
		public void DeleteField(object field)
		{
			 Factory.ExecuteMethod(this, "DeleteField", field);
		}

		#endregion

		#pragma warning restore
	}
}
