using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface CommandBarControls 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class CommandBarControls : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CommandBarControls(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CommandBarControls(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CommandBarControls(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Count"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Parent"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Parent", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		/// <param name="before">optional object before</param>
		/// <param name="temporary">optional object temporary</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before, object temporary)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", new object[]{ type, id, parameter, before, temporary });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id, parameter);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBarControls.Add"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="parameter">optional object parameter</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl Add(object type, object id, object parameter, object before)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "Add", type, id, parameter, before);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CommandBarControl>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
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
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}