using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface Fields 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Fields : COMObject ,IEnumerable<NetOffice.PublisherApi.Field>
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
                    _type = typeof(Fields);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Fields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PublisherApi.Field this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Unlink()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddHorizontalInVertical(NetOffice.PublisherApi.TextRange range, string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text);
			object returnItem = Invoker.MethodReturn(this, "AddHorizontalInVertical", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional object FontSize = 10</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName, object fontSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text, alignment, raise, fontName, fontSize);
			object returnItem = Invoker.MethodReturn(this, "AddPhoneticGuide", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text);
			object returnItem = Invoker.MethodReturn(this, "AddPhoneticGuide", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text, alignment);
			object returnItem = Invoker.MethodReturn(this, "AddPhoneticGuide", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text, alignment, raise);
			object returnItem = Invoker.MethodReturn(this, "AddPhoneticGuide", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange Range</param>
		/// <param name="text">string Text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, text, alignment, raise, fontName);
			object returnItem = Invoker.MethodReturn(this, "AddPhoneticGuide", paramsArray);
			NetOffice.PublisherApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Field.LateBindingApiWrapperType) as NetOffice.PublisherApi.Field;
			return newObject;
		}

		#endregion
       #region IEnumerable<NetOffice.PublisherApi.Field> Member
        
        /// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
        [CustomEnumerator]
       public IEnumerator<NetOffice.PublisherApi.Field> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.PublisherApi.Field item in innerEnumerator)
               yield return item;
       }

       #endregion
   
       #region IEnumerable Members
        
       /// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
       {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i+1];

            foreach (object item in enumeratorObjects)
                yield return item;
       }

       #endregion
       		#pragma warning restore
	}
}