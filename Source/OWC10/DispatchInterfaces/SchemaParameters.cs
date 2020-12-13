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
	/// DispatchInterface SchemaParameters 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SchemaParameters : COMObject, IEnumerableProvider<NetOffice.OWC10Api.SchemaParameter>
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
                    _type = typeof(SchemaParameters);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SchemaParameters(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SchemaParameters(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaParameters(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OWC10Api.SchemaParameter this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaParameter>(this, "Item", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, index);
			}
		}

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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		/// <param name="precision">optional object precision</param>
		/// <param name="direction">optional object direction</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision, object direction)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, new object[]{ name, dataType, size, scale, precision, direction });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, name, dataType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, name, dataType, size);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, name, dataType, size, scale);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="size">optional object size</param>
		/// <param name="scale">optional object scale</param>
		/// <param name="precision">optional object precision</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaParameter Add(string name, object dataType, object size, object scale, object precision)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaParameter>(this, "Add", NetOffice.OWC10Api.SchemaParameter.LateBindingApiWrapperType, new object[]{ name, dataType, size, scale, precision });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.SchemaParameter>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.SchemaParameter>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.SchemaParameter>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaParameter>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<NetOffice.OWC10Api.SchemaParameter> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.SchemaParameter item in innerEnumerator)
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