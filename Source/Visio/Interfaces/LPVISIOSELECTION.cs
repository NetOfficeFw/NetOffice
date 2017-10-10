using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOSELECTION 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class LPVISIOSELECTION : COMObject, IEnumerableProvider<NetOffice.VisioApi.IVShape>
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
                    _type = typeof(LPVISIOSELECTION);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOSELECTION(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOSELECTION(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVShape get_Item16(Int16 index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Item16", NetOffice.VisioApi.IVShape.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Item16
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Item16")]
		public NetOffice.VisioApi.IVShape Item16(Int16 index)
		{
			return get_Item16(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 Count16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ContainingPage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "ContainingMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "ContainingShape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Style
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Style");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string StyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string LineStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string LineStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LineStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FillStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FillStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FillStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FillStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TextStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TextStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVShape this[Int32 index]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 IterationMode
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "IterationMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IterationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ItemStatus(Int32 index)
		{
			return Factory.ExecuteInt16PropertyGet(this, "ItemStatus", index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemStatus
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemStatus")]
		public Int16 ItemStatus(Int32 index)
		{
			return get_ItemStatus(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape PrimaryItem
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "PrimaryItem");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
            }
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ContainingPageID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContainingPageID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ContainingMasterID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContainingMasterID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster DataGraphic
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "DataGraphic");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataGraphic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection SelectionForDragCopy
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SelectionForDragCopy");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Export(string fileName)
		{
			 Factory.ExecuteMethod(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void BringForward()
		{
			 Factory.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void BringToFront()
		{
			 Factory.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SendBackward()
		{
			 Factory.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SendToBack()
		{
			 Factory.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Combine()
		{
			 Factory.ExecuteMethod(this, "Combine");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Fragment()
		{
			 Factory.ExecuteMethod(this, "Fragment");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Intersect()
		{
			 Factory.ExecuteMethod(this, "Intersect");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Subtract()
		{
			 Factory.ExecuteMethod(this, "Subtract");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Union()
		{
			 Factory.ExecuteMethod(this, "Union");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FlipHorizontal()
		{
			 Factory.ExecuteMethod(this, "FlipHorizontal");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FlipVertical()
		{
			 Factory.ExecuteMethod(this, "FlipVertical");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ReverseEnds()
		{
			 Factory.ExecuteMethod(this, "ReverseEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate90()
		{
			 Factory.ExecuteMethod(this, "Rotate90");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_Copy()
		{
			 Factory.ExecuteMethod(this, "old_Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_Cut()
		{
			 Factory.ExecuteMethod(this, "old_Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void VoidDuplicate()
		{
			 Factory.ExecuteMethod(this, "VoidDuplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void VoidGroup()
		{
			 Factory.ExecuteMethod(this, "VoidGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ConvertToGroup()
		{
			 Factory.ExecuteMethod(this, "ConvertToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Ungroup()
		{
			 Factory.ExecuteMethod(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DeselectAll()
		{
			 Factory.ExecuteMethod(this, "DeselectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="selectAction">Int16 selectAction</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction)
		{
			 Factory.ExecuteMethod(this, "Select", sheetObject, selectAction);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Trim()
		{
			 Factory.ExecuteMethod(this, "Trim");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Join()
		{
			 Factory.ExecuteMethod(this, "Join");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FitCurve(Double tolerance, Int16 flags)
		{
			 Factory.ExecuteMethod(this, "FitCurve", tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Layout()
		{
			 Factory.ExecuteMethod(this, "Layout");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true,true);
			lpr8Left = 0;
			lpr8Bottom = 0;
			lpr8Right = 0;
			lpr8Top = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top);
			Invoker.Method(this, "BoundingBox", paramsArray, modifiers);
			lpr8Left = (Double)paramsArray[1];
			lpr8Bottom = (Double)paramsArray[2];
			lpr8Right = (Double)paramsArray[3];
			lpr8Top = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="resultsMaster">optional object resultsMaster</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y, object resultsMaster)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", new object[]{ tolerance, flags, x, y, resultsMaster });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags, x);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags, x, y);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Group()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SwapEnds()
		{
			 Factory.ExecuteMethod(this, "SwapEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddToGroup()
		{
			 Factory.ExecuteMethod(this, "AddToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void RemoveFromGroup()
		{
			 Factory.ExecuteMethod(this, "RemoveFromGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection Duplicate()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "Duplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy(object flags)
		{
			 Factory.ExecuteMethod(this, "Copy", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut(object flags)
		{
			 Factory.ExecuteMethod(this, "Cut", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		/// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Move(Double dx, Double dy, object unitsNameOrCode)
		{
			 Factory.ExecuteMethod(this, "Move", dx, dy, unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Move(Double dx, Double dy)
		{
			 Factory.ExecuteMethod(this, "Move", dx, dy);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			 Factory.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY, pinUnitsNameOrCode });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle)
		{
			 Factory.ExecuteMethod(this, "Rotate", angle);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode)
		{
			 Factory.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards)
		{
			 Factory.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode, blastGuards);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType)
		{
			 Factory.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode, blastGuards, rotationType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX)
		{
			 Factory.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY)
		{
			 Factory.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical, object glueToGuide)
		{
			 Factory.ExecuteMethod(this, "Align", alignHorizontal, alignVertical, glueToGuide);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical)
		{
			 Factory.ExecuteMethod(this, "Align", alignHorizontal, alignVertical);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute, object glueToGuide)
		{
			 Factory.ExecuteMethod(this, "Distribute", distribute, glueToGuide);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute)
		{
			 Factory.ExecuteMethod(this, "Distribute", distribute);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void UpdateAlignmentBox()
		{
			 Factory.ExecuteMethod(this, "UpdateAlignmentBox");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distance">Double distance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Offset(Double distance)
		{
			 Factory.ExecuteMethod(this, "Offset", distance);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ConnectShapes()
		{
			 Factory.ExecuteMethod(this, "ConnectShapes");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			 Factory.ExecuteMethod(this, "Flip", new object[]{ flipDirection, flipType, blastGuards, pinX, pinY, pinUnitsNameOrCode });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection)
		{
			 Factory.ExecuteMethod(this, "Flip", flipDirection);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType)
		{
			 Factory.ExecuteMethod(this, "Flip", flipDirection, flipType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards)
		{
			 Factory.ExecuteMethod(this, "Flip", flipDirection, flipType, blastGuards);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX)
		{
			 Factory.ExecuteMethod(this, "Flip", flipDirection, flipType, blastGuards, pinX);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY)
		{
			 Factory.ExecuteMethod(this, "Flip", new object[]{ flipDirection, flipType, blastGuards, pinX, pinY });
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 dataRowID, object applyDataGraphicAfterLink)
		{
			 Factory.ExecuteMethod(this, "LinkToData", dataRecordsetID, dataRowID, applyDataGraphicAfterLink);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 dataRowID)
		{
			 Factory.ExecuteMethod(this, "LinkToData", dataRecordsetID, dataRowID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void BreakLinkToData(Int32 dataRecordsetID)
		{
			 Factory.ExecuteMethod(this, "BreakLinkToData", dataRecordsetID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void GetIDs(out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)shapeIDs);
			Invoker.Method(this, "GetIDs", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="columnNames">String[] columnNames</param>
		/// <param name="autoLinkFieldTypes">Int32[] autoLinkFieldTypes</param>
		/// <param name="fieldNames">String[] fieldNames</param>
		/// <param name="autoLinkBehavior">Int32 autoLinkBehavior</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 AutomaticLink(Int32 dataRecordsetID, String[] columnNames, Int32[] autoLinkFieldTypes, String[] fieldNames, Int32 autoLinkBehavior, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)columnNames, (object)autoLinkFieldTypes, (object)fieldNames, autoLinkBehavior, (object)shapeIDs);
			object returnItem = Invoker.MethodReturn(this, "AutomaticLink", paramsArray);
			shapeIDs = (Int32[])paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical</param>
		/// <param name="spaceHorizontal">Double spaceHorizontal</param>
		/// <param name="spaceVertical">Double spaceVertical</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			 Factory.ExecuteMethod(this, "LayoutIncremental", new object[]{ alignOrSpace, alignHorizontal, alignVertical, spaceHorizontal, spaceVertical, unitCode });
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			 Factory.ExecuteMethod(this, "LayoutChangeDirection", direction);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void AvoidPageBreaks()
		{
			 Factory.ExecuteMethod(this, "AvoidPageBreaks");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
		/// <param name="distance">Double distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			 Factory.ExecuteMethod(this, "Resize", direction, distance, unitCode);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void AddToContainers()
		{
			 Factory.ExecuteMethod(this, "AddToContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void RemoveFromContainers()
		{
			 Factory.ExecuteMethod(this, "RemoveFromContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop, newShape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="delFlags">Int32 delFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void DeleteEx(Int32 delFlags)
		{
			 Factory.ExecuteMethod(this, "DeleteEx", delFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetContainers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetCallouts", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] MemberOfContainersUnion()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersUnion", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] MemberOfContainersIntersection()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersIntersection", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		/// <param name="formatValue">optional object FormatValue = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType, object formatValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType, formatValue);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop, replaceFlags);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersion("Visio", 15, 16)]
		public void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
		{
			 Factory.ExecuteMethod(this, "SetQuickStyle", new object[]{ lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVShape>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVShape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVShape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVShape>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.VisioApi.IVShape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVShape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}