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
    /// DispatchInterface CommandBarControls 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862747.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
    public class CommandBarControls : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.CommandBarControls
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
                    _type = typeof(CommandBarControls);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public CommandBarControls() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860596.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.CommandBarControl this[object index]
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OfficeApi.CommandBarControl>(this, "Item", index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860798.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBar Parent
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Parent", typeof(NetOffice.OfficeApi.CommandBar));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="parameter">optional object parameter</param>
        /// <param name="before">optional object before</param>
        /// <param name="temporary">optional object temporary</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before, object temporary)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", new object[] { type, id, parameter, before, temporary });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBarControl Add()
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBarControl Add(object type)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBarControl Add(object type, object id)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="parameter">optional object parameter</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id, parameter);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861771.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="parameter">optional object parameter</param>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id, parameter, before);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CommandBarControl>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.CommandBarControl> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.CommandBarControl item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
