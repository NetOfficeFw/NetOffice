using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// Interface LPVISIOCONTAINERPROPERTIES 
	/// SupportByVersion Visio, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIOCONTAINERPROPERTIES : COMObject
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
                    _type = typeof(LPVISIOCONTAINERPROPERTIES);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOCONTAINERPROPERTIES(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCONTAINERPROPERTIES(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerTypes ContainerType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainerType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisContainerTypes)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisListAlignment ListAlignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListAlignment", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisListAlignment)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ListAlignment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisListDirection ListDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisListDirection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ListDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool LockMembership
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockMembership", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockMembership", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerAutoResize ResizeAsNeeded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ResizeAsNeeded", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisContainerAutoResize)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ResizeAsNeeded", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape OverlappedList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OverlappedList", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OverlappedList", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 ContainerStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainerStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ContainerStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 HeadingStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeadingStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeadingStyle", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void Disband()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Disband", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void FitToContents()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FitToContents", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes MarginUnits</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Double GetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(marginUnits);
			object returnItem = Invoker.MethodReturn(this, "GetMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="marginUnits">NetOffice.VisioApi.Enums.VisUnitCodes MarginUnits</param>
		/// <param name="marginSize">Double MarginSize</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void SetMargin(NetOffice.VisioApi.Enums.VisUnitCodes marginUnits, Double marginSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(marginUnits, marginSize);
			Invoker.Method(this, "SetMargin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes SpacingUnits</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Double GetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(spacingUnits);
			object returnItem = Invoker.MethodReturn(this, "GetListSpacing", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="spacingUnits">NetOffice.VisioApi.Enums.VisUnitCodes SpacingUnits</param>
		/// <param name="spacingSize">Double SpacingSize</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void SetListSpacing(NetOffice.VisioApi.Enums.VisUnitCodes spacingUnits, Double spacingSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(spacingUnits, spacingSize);
			Invoker.Method(this, "SetListSpacing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToInsert">object ObjectToInsert</param>
		/// <param name="position">Int32 Position</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void InsertListMember(object objectToInsert, Int32 position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToInsert, position);
			Invoker.Method(this, "InsertListMember", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shapeMember">NetOffice.VisioApi.IVShape ShapeMember</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 GetListMemberPosition(NetOffice.VisioApi.IVShape shapeMember)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shapeMember);
			object returnItem = Invoker.MethodReturn(this, "GetListMemberPosition", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shape">NetOffice.VisioApi.IVShape Shape</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisContainerMemberState GetMemberState(NetOffice.VisioApi.IVShape shape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape);
			object returnItem = Invoker.MethodReturn(this, "GetMemberState", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.VisioApi.Enums.VisContainerMemberState)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToRemove">object ObjectToRemove</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void RemoveMember(object objectToRemove)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToRemove);
			Invoker.Method(this, "RemoveMember", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToReorder">object ObjectToReorder</param>
		/// <param name="position">Int32 Position</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void ReorderListMember(object objectToReorder, Int32 position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToReorder, position);
			Invoker.Method(this, "ReorderListMember", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GetListMembers()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetListMembers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="containerFlags">Int32 ContainerFlags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GetMemberShapes(Int32 containerFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(containerFlags);
			object returnItem = (object)Invoker.MethodReturn(this, "GetMemberShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pObjectToAdd">object pObjectToAdd</param>
		/// <param name="addOptions">NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void AddMember(object pObjectToAdd, NetOffice.VisioApi.Enums.VisMemberAddOptions addOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pObjectToAdd, addOptions);
			Invoker.Method(this, "AddMember", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection Direction</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void RotateFlipList(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "RotateFlipList", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}