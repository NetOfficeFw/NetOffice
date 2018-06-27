using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.DAOApi.Behind
{
    /// <summary>
    /// DispatchInterface IndexFields 
    /// SupportByVersion DAO, 3.6,12.0
    /// </summary>
    public class IndexFields : _DynaCollection, NetOffice.DAOApi.IndexFields
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
                    _contractType = typeof(NetOffice.DAOApi.IndexFields);
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
                    _type = typeof(IndexFields);
                return _type;
            }
        }

        #endregion

        #region Ctor    

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IndexFields() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion DAO 3.6, 12.0
        /// Get
        /// </summary>
        /// <param name="item">optional object item</param>
        [SupportByVersion("DAO", 3.6, 12.0)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public object this[object item]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Item", item);
            }
        }

        #endregion

        #pragma warning restore
    }
}
