using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface SchemaRowsources 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SchemaRowsources : COMObject, IEnumerableProvider<NetOffice.OWC10Api.SchemaRowsource>
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
                    _type = typeof(SchemaRowsources);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SchemaRowsources(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SchemaRowsources(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRowsources(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OWC10Api.SchemaRowsource this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRowsource>(this, "Item", NetOffice.OWC10Api.SchemaRowsource.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="rowsourceType">NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType</param>
		/// <param name="commandText">optional object commandText</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsource Add(string name, NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType, object commandText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRowsource>(this, "Add", NetOffice.OWC10Api.SchemaRowsource.LateBindingApiWrapperType, name, rowsourceType, commandText);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="rowsourceType">NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsource Add(string name, NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRowsource>(this, "Add", NetOffice.OWC10Api.SchemaRowsource.LateBindingApiWrapperType, name, rowsourceType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="rowsourceType">NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType</param>
		/// <param name="commandText">optional object commandText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsource AddNew(string name, NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType, object commandText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRowsource>(this, "AddNew", NetOffice.OWC10Api.SchemaRowsource.LateBindingApiWrapperType, name, rowsourceType, commandText);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="rowsourceType">NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsource AddNew(string name, NetOffice.OWC10Api.Enums.DscRowsourceTypeEnum rowsourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRowsource>(this, "AddNew", NetOffice.OWC10Api.SchemaRowsource.LateBindingApiWrapperType, name, rowsourceType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		public void Delete(object index)
		{
			 Factory.ExecuteMethod(this, "Delete", index);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.SchemaRowsource>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.SchemaRowsource>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.SchemaRowsource>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaRowsource>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<NetOffice.OWC10Api.SchemaRowsource> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.SchemaRowsource item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}