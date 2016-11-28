using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface GradientStops 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861159.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class GradientStops : _IMsoDispObj ,IEnumerable<NetOffice.OfficeApi.GradientStop>
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
                    _type = typeof(GradientStops);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public GradientStops(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.GradientStop this[Int32 index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OfficeApi.GradientStop newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.GradientStop.LateBindingApiWrapperType) as NetOffice.OfficeApi.GradientStop;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864855.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx
		/// </summary>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Delete(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position, object transparency, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position, transparency, index);
			Invoker.Method(this, "Insert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position);
			Invoker.Method(this, "Insert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position, object transparency)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position, transparency);
			Invoker.Method(this, "Insert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		/// <param name="brightness">optional Single Brightness = 0</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency, object index, object brightness)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position, transparency, index, brightness);
			Invoker.Method(this, "Insert2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position);
			Invoker.Method(this, "Insert2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position, transparency);
			Invoker.Method(this, "Insert2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx
		/// </summary>
		/// <param name="rGB">Int32 RGB</param>
		/// <param name="position">Single Position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rGB, position, transparency, index);
			Invoker.Method(this, "Insert2", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.GradientStop> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.GradientStop> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.GradientStop item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Office, 12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}