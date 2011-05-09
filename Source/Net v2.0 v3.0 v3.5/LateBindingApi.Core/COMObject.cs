using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// represents a managed COM proxy 
    /// </summary>
    public class COMObject : IDisposable 
    {
        #region Fields

        protected internal COMObject            _parentObject;
        protected internal Type                 _instanceType;
        protected internal object               _underlyingObject;

        protected internal volatile bool        _isCurrentlyDisposing;
        protected internal volatile bool        _isDisposed;
        protected internal List<COMObject>      _listChildObjects    = new List<COMObject>();
       
        #endregion

        #region Construction

        public COMObject(COMObject replacedObject)
        {
            // copy proxy
            _underlyingObject = replacedObject.UnderlyingObject;
            _parentObject     = replacedObject.ParentObject;
            _instanceType     = replacedObject.InstanceType;

            // copy childs
            foreach (IObject item in replacedObject.ListChildObjects)
                AddChildObject(item);

            // remove old object from parent chain
            if (null != replacedObject.ParentObject)
            {
                COMObject parentObject = replacedObject.ParentObject;
                parentObject.RemoveChildObject(replacedObject);
                
                // add himself as child to parent object
                parentObject.AddChildObject(this);
            }

            Factory.RemoveObjectFromList(replacedObject);
            Factory.AddObjectToList(this);
        }
        
        public COMObject(object comProxy)
        {
            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            Factory.AddObjectToList(this);
        }

        public COMObject(COMObject parentObject, object comProxy)
        {
            _parentObject = parentObject;
            _underlyingObject = comProxy;
            _instanceType = comProxy.GetType();

            if (null!= parentObject)
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        public COMObject(COMObject parentObject, object comProxy, Type comProxyType)
        {
            _parentObject = parentObject;
            _underlyingObject = comProxy;
            _instanceType = comProxyType;

            if(null!=parentObject)
                _parentObject.AddChildObject(this);

            Factory.AddObjectToList(this);
        }

        public COMObject()
        {
            Factory.AddObjectToList(this);
        }

        #endregion

        #region COMObject Properties

        /// <summary>
        /// returns the native wrapped proxy
        /// </summary>
        public object UnderlyingObject
        {
            get
            {
                return _underlyingObject;
            }
        }

        /// <summary>
        /// returns class name of native wrapped proxy
        /// </summary>
        public string UnderlyingTypeName
        {
            get
            {
                return TypeDescriptor.GetClassName(_underlyingObject);
            }
        }

        /// <summary>
        /// returns instance is diposed means unusable
        /// </summary>
        public bool IsDisposed
        {
            get 
            {
                return _isDisposed;
            }
        }

        /// <summary>
        /// returns Type of native proxy
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Type InstanceType
        {
            get
            {
                return _instanceType;
            }
        }

        /// <summary>
        /// returns parent proxy object
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public COMObject ParentObject
        {
            get
            {
                return _parentObject;
            }
            set
            {
                _parentObject = value;
            }
        }

        /// <summary>
        /// returns instance is currently in diposing progress
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool IsCurrentlyDisposing
        {
            get
            {
                return _isCurrentlyDisposing;
            }
        }
        
        /// <summary>
        /// returns instance export events
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool IsEventBind
        {
            get 
            {
                return (null != (this as IEventBinding));
            }
        }

        /// <summary>
        /// returns event bridge is advised
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get
            {
                IEventBinding bindInfo = this as IEventBinding;
                if (null != bindInfo)
                    return bindInfo.EventBridgeInitialized;
                else
                    return false;
            }
        }

        /// <summary>
        ///  retuns instance has one or more event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients
        {
            get
            {
                IEventBinding bindInfo = this as IEventBinding;
                if(null!=bindInfo)
                    return bindInfo.HasEventRecipients;
                else
                    return false;
            }
        }
        /// <summary>
        ///  child instance List
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        internal List<COMObject> ListChildObjects
        {
            get
            {
                return _listChildObjects;
            }
        }

        #endregion

        #region COMObject Methods

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void CreateFromProgId(string progId)
        {
            _instanceType = System.Type.GetTypeFromProgID(progId);
            if (null == _instanceType)
                throw (new ArgumentException("progId not found. " + progId));

            _underlyingObject = Activator.CreateInstance(_instanceType);
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void AddChildObject(IObject childObject)
        {
            if (childObject is COMObject)
                _listChildObjects.Add((COMObject)childObject);

            throw new ArgumentException("childObject is an unkown type.");
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void AddChildObject(COMObject childObject)
        {
            _listChildObjects.Add(childObject);
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void RemoveChildObject(COMObject childObject)
        {
            _listChildObjects.Remove(childObject);
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void ReleaseCOMProxy()
        {
            // release himself from COM Runtime System
            if (null != _underlyingObject)
            {
                if (_underlyingObject is ICustomAdapter)
                {
                    // enumerator
                    ICustomAdapter adapter = (ICustomAdapter)_underlyingObject;
                    Marshal.ReleaseComObject(adapter.GetUnderlyingObject());
                }
                else
                {
                    Marshal.ReleaseComObject(_underlyingObject);
                }
                Factory.RemoveObjectFromList(this);
                _underlyingObject = null;
            }
        }

     
        #endregion

        #region IDisposable Members

        /// <summary>
        /// dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventProxies">dispose event exported proxies with one or more event recipients</param>
        public void Dispose(bool disposeEventBinding)
        {
            // set disposed flag
            _isCurrentlyDisposing = true;

            // in case of object implements also event binding we dispose them
            IEventBinding eventBind = this as IEventBinding;
            if (disposeEventBinding)
            {
                if (null != eventBind)
                    eventBind.DisposeSinkHelper();
            }
            else
            {
                if( (null != eventBind) && (!eventBind.EventBridgeInitialized) )
                    eventBind.DisposeSinkHelper();
            }
 
            // child proxy dispose
            DisposeChildInstances(disposeEventBinding);
            
            // remove himself from parent childlist
            if (null != _parentObject)
            {
                _parentObject.RemoveChildObject(this);
                _parentObject = null;
            }

            // release proxy
            ReleaseCOMProxy();

            _isDisposed = true;
            _isCurrentlyDisposing = false;
        }

        /// <summary>
        /// dispose instance and all child instances
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// dispose all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        public void DisposeChildInstances(bool disposeEventBinding)
        {
            // release all childs and clear list
            foreach (COMObject itemObject in _listChildObjects)
            {
                itemObject.ParentObject = null;
                itemObject.Dispose(disposeEventBinding);
            }
            _listChildObjects.Clear();
        }

        /// <summary>
        /// dispose all child instances
        /// </summary>
        public void DisposeChildInstances()
        {
            // release all childs and clear list
            foreach (COMObject itemObject in _listChildObjects)
            {
                itemObject.ParentObject = null;
                itemObject.Dispose();
            }
            _listChildObjects.Clear();
        }

        #endregion
    }
}
