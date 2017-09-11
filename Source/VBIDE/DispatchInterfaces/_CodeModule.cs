using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface _CodeModule 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CodeModule : COMObject
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
                    _type = typeof(_CodeModule);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CodeModule(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CodeModule(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodeModule(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBComponent Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBComponent>(this, "Parent", NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Lines(Int32 startLine, Int32 count)
		{
			return Factory.ExecuteStringPropertyGet(this, "Lines", startLine, count);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_Lines
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_Lines")]
		public string Lines(Int32 startLine, Int32 count)
		{
			return get_Lines(startLine, count);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 CountOfLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CountOfLines");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcStartLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcStartLine
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_ProcStartLine")]
		public Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcStartLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcCountLines", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcCountLines
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_ProcCountLines")]
		public Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcCountLines(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcBodyLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcBodyLine
		/// </summary>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_ProcBodyLine")]
		public Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcBodyLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="line">Int32 line</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			procKind = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(line, procKind);
			object returnItem = Invoker.PropertyGet(this, "ProcOfLine", paramsArray, modifiers);
			procKind = (NetOffice.VBIDEApi.Enums.vbext_ProcKind)paramsArray[1];
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ProcOfLine
		/// </summary>
		/// <param name="line">Int32 line</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_ProcOfLine")]
		public string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcOfLine(line, out procKind);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 CountOfDeclarationLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CountOfDeclarationLines");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodePane CodePane
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePane>(this, "CodePane", NetOffice.VBIDEApi.CodePane.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="_string">string string</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void AddFromString(string _string)
		{
			 Factory.ExecuteMethod(this, "AddFromString", _string);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void AddFromFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "AddFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void InsertLines(Int32 line, string _string)
		{
			 Factory.ExecuteMethod(this, "InsertLines", line, _string);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void DeleteLines(Int32 startLine, object count)
		{
			 Factory.ExecuteMethod(this, "DeleteLines", startLine, count);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		[CustomMethod]
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void DeleteLines(Int32 startLine)
		{
			 Factory.ExecuteMethod(this, "DeleteLines", startLine);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ReplaceLine(Int32 line, string _string)
		{
			 Factory.ExecuteMethod(this, "ReplaceLine", line, _string);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="eventName">string eventName</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 CreateEventProc(string eventName, string objectName)
		{
			return Factory.ExecuteInt32MethodGet(this, "CreateEventProc", eventName, objectName);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		/// <param name="patternSearch">optional bool PatternSearch = false</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, patternSearch });
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[CustomMethod]
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn });
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		[CustomMethod]
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord });
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		[CustomMethod]
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase });
		}

		#endregion

		#pragma warning restore
	}
}
