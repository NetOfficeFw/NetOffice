using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface AddIn
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class AddIn : COMObject, NetOffice.VBIDEApi.AddIn
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
                    _type = typeof(AddIn);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public AddIn() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string Description
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Description");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Description", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", typeof(NetOffice.VBIDEApi.VBE));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.Addins Collection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Addins>(this, "Collection", typeof(NetOffice.VBIDEApi.Addins));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string ProgId
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ProgId");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string Guid
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Guid");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public bool Connect
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Connect");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Connect", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3), ProxyResult]
        public object Object
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Object");
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "Object", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
