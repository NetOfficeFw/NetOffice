using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// Interface ADOStreamConstruction 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsInterface)]
 	public class ADOStreamConstruction : COMObject, NetOffice.ADODBApi.ADOStreamConstruction
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
                    _type = typeof(ADOStreamConstruction);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ADOStreamConstruction() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
		public virtual object Stream
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Stream");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Stream", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

