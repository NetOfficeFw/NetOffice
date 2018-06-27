using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface Property
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Property : COMObject, NetOffice.VBIDEApi.Property
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
                    _contractType = typeof(NetOffice.VBIDEApi.Property);
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
                    _type = typeof(Property);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Property() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual object Value
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_IndexedValue(object index1, object index2, object index3, object index4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2, index3, index4);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_IndexedValue(object index1, object index2, object index3, object index4, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "IndexedValue", new object[] { index1, index2, index3, index4, value });
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        public virtual object IndexedValue(object index1, object index2, object index3, object index4)
        {
            return get_IndexedValue(index1, index2, index3, index4);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_IndexedValue(object index1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IndexedValue", index1);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_IndexedValue(object index1, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "IndexedValue", index1, value);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        public virtual object IndexedValue(object index1)
        {
            return get_IndexedValue(index1);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_IndexedValue(object index1, object index2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_IndexedValue(object index1, object index2, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "IndexedValue", index1, index2, value);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        public virtual object IndexedValue(object index1, object index2)
        {
            return get_IndexedValue(index1, index2);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_IndexedValue(object index1, object index2, object index3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IndexedValue", index1, index2, index3);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_IndexedValue(object index1, object index2, object index3, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "IndexedValue", index1, index2, index3, value);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        public virtual object IndexedValue(object index1, object index2, object index3)
        {
            return get_IndexedValue(index1, index2, index3);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int16 NumIndices
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "NumIndices");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.VBIDEApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VBIDEApi.Application>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.VBIDEApi.Properties Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Properties>(this, "Parent", typeof(NetOffice.VBIDEApi.Properties));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Properties Collection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Properties>(this, "Collection", typeof(NetOffice.VBIDEApi.Properties));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3), ProxyResult]
        public virtual object Object
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Object");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Object", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
