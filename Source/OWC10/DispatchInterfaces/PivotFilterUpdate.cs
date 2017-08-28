using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotFilterUpdate 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotFilterUpdate : COMObject
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
                    _type = typeof(PivotFilterUpdate);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotFilterUpdate(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotFilterUpdate(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilterUpdate(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum get_StateOf(NetOffice.OWC10Api.PivotMember member)
		{
			return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum>(this, "StateOf", member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_StateOf
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1), Redirect("get_StateOf")]
		public NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum StateOf(NetOffice.OWC10Api.PivotMember member)
		{
			return get_StateOf(member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsDirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDirty");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		public void Click(NetOffice.OWC10Api.PivotMember member)
		{
			 Factory.ExecuteMethod(this, "Click", member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		/// <param name="oldMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState</param>
		/// <param name="newMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState</param>
		[SupportByVersion("OWC10", 1)]
		public void ClickFromTo(NetOffice.OWC10Api.PivotMember member, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState)
		{
			 Factory.ExecuteMethod(this, "ClickFromTo", member, oldMemberState, newMemberState);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Apply()
		{
			 Factory.ExecuteMethod(this, "Apply");
		}

		#endregion

		#pragma warning restore
	}
}
