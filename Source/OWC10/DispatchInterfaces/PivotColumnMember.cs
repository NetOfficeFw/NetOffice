using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotColumnMember 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotColumnMember : PivotAxisMember
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
                    _type = typeof(PivotColumnMember);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotColumnMember(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotColumnMember(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotColumnMember(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMembers ChildColumnMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMembers>(this, "ChildColumnMembers", NetOffice.OWC10Api.PivotColumnMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMember ParentColumnMember
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "ParentColumnMember", NetOffice.OWC10Api.PivotColumnMember.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotColumnMember get_FindColumnMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "FindColumnMember", NetOffice.OWC10Api.PivotColumnMember.LateBindingApiWrapperType, path, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindColumnMember
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindColumnMember")]
		public NetOffice.OWC10Api.PivotColumnMember FindColumnMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format)
		{
			return get_FindColumnMember(path, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotColumnMember TotalColumnMember
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotColumnMember>(this, "TotalColumnMember", NetOffice.OWC10Api.PivotColumnMember.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailLeftOffset
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailLeftOffset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailLeftOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DetailsExpanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DetailsExpanded");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailLeft">Int32 detailLeft</param>
		/// <param name="detailLeftOffset">Int32 detailLeftOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		public void MoveDetailLeft(Int32 detailLeft, Int32 detailLeftOffset, object update)
		{
			 Factory.ExecuteMethod(this, "MoveDetailLeft", detailLeft, detailLeftOffset, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailLeft">Int32 detailLeft</param>
		/// <param name="detailLeftOffset">Int32 detailLeftOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void MoveDetailLeft(Int32 detailLeft, Int32 detailLeftOffset)
		{
			 Factory.ExecuteMethod(this, "MoveDetailLeft", detailLeft, detailLeftOffset);
		}

		#endregion

		#pragma warning restore
	}
}
