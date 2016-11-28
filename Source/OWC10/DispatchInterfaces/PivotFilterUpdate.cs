using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface PivotFilterUpdate 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PivotFilterUpdate : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="member">NetOffice.OWC10Api.PivotMember Member</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum get_StateOf(NetOffice.OWC10Api.PivotMember member)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(member);
			object returnItem = Invoker.PropertyGet(this, "StateOf", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_StateOf
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember Member</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum StateOf(NetOffice.OWC10Api.PivotMember member)
		{
			return get_StateOf(member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool IsDirty
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDirty", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember Member</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Click(NetOffice.OWC10Api.PivotMember member)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(member);
			Invoker.Method(this, "Click", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember Member</param>
		/// <param name="oldMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum OldMemberState</param>
		/// <param name="newMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum NewMemberState</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ClickFromTo(NetOffice.OWC10Api.PivotMember member, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(member, oldMemberState, newMemberState);
			Invoker.Method(this, "ClickFromTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Apply()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Apply", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}