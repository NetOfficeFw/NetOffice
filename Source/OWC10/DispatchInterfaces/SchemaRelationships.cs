using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface SchemaRelationships 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SchemaRelationships : COMObject ,IEnumerable<NetOffice.OWC10Api.SchemaRelationship>
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OWC10Api.SchemaRelationship this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OWC10Api.SchemaRelationship newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaRelationship;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="manySchemaRowsource">string ManySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string OneSchemaRowsource</param>
		/// <param name="manySchemaField">string ManySchemaField</param>
		/// <param name="oneSchemaField">string OneSchemaField</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationship Add(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OWC10Api.SchemaRelationship newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaRelationship;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="manySchemaRowsource">string ManySchemaRowsource</param>
		/// <param name="oneSchemaRowsource">string OneSchemaRowsource</param>
		/// <param name="manySchemaField">string ManySchemaField</param>
		/// <param name="oneSchemaField">string OneSchemaField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationship AddNew(string name, string manySchemaRowsource, string oneSchemaRowsource, string manySchemaField, string oneSchemaField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, manySchemaRowsource, oneSchemaRowsource, manySchemaField, oneSchemaField);
			object returnItem = Invoker.MethodReturn(this, "AddNew", paramsArray);
			NetOffice.OWC10Api.SchemaRelationship newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.SchemaRelationship.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaRelationship;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Delete", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OWC10Api.SchemaRelationship> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<NetOffice.OWC10Api.SchemaRelationship> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OWC10Api.SchemaRelationship item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}