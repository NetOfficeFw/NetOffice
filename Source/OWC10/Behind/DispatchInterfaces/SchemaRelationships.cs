using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface SchemaRelationships 
	/// SupportByVersion OWC10, 1
	/// </summary>
	public class SchemaRelationships : COMObject, NetOffice.OWC10Api.SchemaRelationships
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
                    _contractType = typeof(NetOffice.OWC10Api.SchemaRelationships);
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
                    _type = typeof(SchemaRelationships);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SchemaRelationships() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.OWC10Api.SchemaRelationship this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRelationship>(this, "Item", typeof(NetOffice.OWC10Api.SchemaRelationship), index);
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
		public virtual NetOffice.OWC10Api.SchemaRelationship Add(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRelationship>(this, "Add", typeof(NetOffice.OWC10Api.SchemaRelationship), new object[]{ name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField });
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
		public virtual NetOffice.OWC10Api.SchemaRelationship AddNew(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.SchemaRelationship>(this, "AddNew", typeof(NetOffice.OWC10Api.SchemaRelationship), new object[]{ name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Delete(object index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index);
		}

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.SchemaRelationship> Member

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual IEnumerator<NetOffice.OWC10Api.SchemaRelationship> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}


