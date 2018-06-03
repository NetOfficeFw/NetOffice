using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface PictureEffects 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864059.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
    public class PictureEffects : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.PictureEffects
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
                    _type = typeof(PictureEffects);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PictureEffects() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.PictureEffect this[Int32 index]
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureEffect>(this, "Item", typeof(NetOffice.OfficeApi.PictureEffect), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861170.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861830.aspx </remarks>
        /// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
        /// <param name="position">optional Int32 Position = -1</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType, object position)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PictureEffect>(this, "Insert", typeof(NetOffice.OfficeApi.PictureEffect), effectType, position);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861830.aspx </remarks>
        /// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PictureEffect>(this, "Insert", typeof(NetOffice.OfficeApi.PictureEffect), effectType);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862552.aspx </remarks>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Delete(object index)
        {
            Factory.ExecuteMethod(this, "Delete", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862552.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PictureEffect>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.PictureEffect> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.PictureEffect item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
