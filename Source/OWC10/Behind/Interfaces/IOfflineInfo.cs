using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IOfflineInfo 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IOfflineInfo : COMObject, NetOffice.OWC10Api.IOfflineInfo
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
                    _contractType = typeof(NetOffice.OWC10Api.IOfflineInfo);
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
                    _type = typeof(IOfflineInfo);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IOfflineInfo() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pwzServerFilter">string pwzServerFilter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PutServerFilter(string pwzUrl, string pwzServerFilter)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PutServerFilter", pwzUrl, pwzServerFilter);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pwzServerFilter">string pwzServerFilter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetServerFilter(string pwzUrl, string pwzServerFilter)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetServerFilter", pwzUrl, pwzServerFilter);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pwzUrl">string pwzUrl</param>
		/// <param name="pfSubscribed">Int32 pfSubscribed</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetIsPageSubscribed(string pwzUrl, Int32 pfSubscribed)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetIsPageSubscribed", pwzUrl, pfSubscribed);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstrPath">string pbstrPath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetOfflineXMLFileLocation(string pbstrPath)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetOfflineXMLFileLocation", pbstrPath);
		}

		#endregion

		#pragma warning restore
	}
}

