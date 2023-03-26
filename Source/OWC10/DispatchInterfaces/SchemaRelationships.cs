using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface SchemaRelationships 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SchemaRelationships : COMObject, IEnumerable<NetOffice.OWC10Api.SchemaRelationship>
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
                    _type = typeof(SchemaRelationships);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SchemaRelationships(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SchemaRelationships(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SchemaRelationships(string progId) : base(progId)
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
		public NetOffice.OWC10Api.SchemaRelationship this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRelationship>(this, "Item", NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="manySchemaRowsource">string manySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string oneSchemaRowsource</param>
		/// <param name="manySchemaField">string manySchemaField</param>
		/// <param name="oneSchemaField">string oneSchemaField</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationship Add(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRelationship>(this, "Add", NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType, new object[]{ name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="manySchemaRowsource">string manySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string oneSchemaRowsource</param>
		/// <param name="manySchemaField">string manySchemaField</param>
		/// <param name="oneSchemaField">string oneSchemaField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationship AddNew(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRelationship>(this, "AddNew", NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType, new object[]{ name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField });
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

        #region IEnumerable<NetOffice.OWC10Api.SchemaRelationship> Member

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<NetOffice.OWC10Api.SchemaRelationship> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.SchemaRelationship item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

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
