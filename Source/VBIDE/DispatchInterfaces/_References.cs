using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.VBIDEApi
{
	///<summary>
	/// DispatchInterface _References 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _References : COMObject ,IEnumerable<NetOffice.VBIDEApi.Reference>
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
                    _type = typeof(_References);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _References(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _References(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBProject Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.VBIDEApi.VBProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.VBIDEApi.Reference this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.VBIDEApi.Reference newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.VBIDEApi.Reference.LateBindingApiWrapperType) as NetOffice.VBIDEApi.Reference;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="guid">string Guid</param>
		/// <param name="major">Int32 Major</param>
		/// <param name="minor">Int32 Minor</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Reference AddFromGuid(string guid, Int32 major, Int32 minor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(guid, major, minor);
			object returnItem = Invoker.MethodReturn(this, "AddFromGuid", paramsArray);
			NetOffice.VBIDEApi.Reference newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.VBIDEApi.Reference.LateBindingApiWrapperType) as NetOffice.VBIDEApi.Reference;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Reference AddFromFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "AddFromFile", paramsArray);
			NetOffice.VBIDEApi.Reference newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.VBIDEApi.Reference.LateBindingApiWrapperType) as NetOffice.VBIDEApi.Reference;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="reference">NetOffice.VBIDEApi.Reference Reference</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void Remove(NetOffice.VBIDEApi.Reference reference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reference);
			Invoker.Method(this, "Remove", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.VBIDEApi.Reference> Member
        
        /// <summary>
		/// SupportByVersionAttribute VBIDE, 12,14,5.3
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
       public IEnumerator<NetOffice.VBIDEApi.Reference> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VBIDEApi.Reference item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute VBIDE, 12,14,5.3
		/// </summary>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}