using System;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Security;
using System.Security.Permissions;

namespace NOTools.ComponentModel
{
    /// <summary>Stellt eine generische Auflistung bereit, die eine Datenbindung unterstützt.</summary>
    /// <typeparam name="T">Der Typ der Elemente in der Liste.</typeparam>
    [HostProtection(SecurityAction.LinkDemand, SharedState = true)]
    [Serializable]
    public abstract class BindingList<T> : NOTools.ComponentModel.Collection<T>, IBindingList, IList, ICollection, IEnumerable, ICancelAddNew, IRaiseItemChangedEvents
    {
        #region Fields

        private int _addNewPos = -1;

        private bool _raiseListChangedEvents = true;
      
        private bool _raiseItemChangedEvents;
      
        [NonSerialized]
        private PropertyDescriptorCollection _itemTypeProperties;
      
        [NonSerialized]
        private PropertyChangedEventHandler _propertyChangedEventHandler;
       
        [NonSerialized]
        private AddingNewEventHandler _onAddingNew;
      
        [NonSerialized]
        private ListChangedEventHandler _onListChanged;
     
        [NonSerialized]
        private int _lastChangeIndex = -1;
       
        private bool _allowNew = true;
        
        private bool _allowEdit = true;
       
        private bool _allowRemove = true;
        
        private bool _userSetAllowNew;

        #endregion

        #region Ctor

        /// <summary>Initialisiert mithilfe von Standardwerten eine neue Instanz der <see cref="T:System.ComponentModel.BindingList`1" />-Klasse </summary>
        public BindingList()
        {
            Initialize();
        }

        /// <summary>Initialisiert eine neue Instanz der <see cref="T:System.ComponentModel.BindingList`1" />-Klasse mit der angegebenen Liste.</summary>
        /// <param name="list">Eine <see cref="T:System.Collections.Generic.IList`1" /> von Elementen, die in <see cref="T:System.ComponentModel.BindingList`1" /> enthalten sein sollen.</param>
        public BindingList(IList<T> list) : base(list)
        {
            Initialize();
        }

        private void Initialize()
        {
            //_allowNew = this.ItemTypeHasDefaultConstructor;
            if (typeof(INotifyPropertyChanged).IsAssignableFrom(typeof(T)))
            {
                _raiseItemChangedEvents = true;
                foreach (T current in base.Items)
                    this.HookPropertyChanged(current);
            }
        }

        #endregion

        #region Events

        /// <summary>Tritt ein, bevor der Liste ein Element hinzugefügt wird.</summary>
        public event AddingNewEventHandler AddingNew
        {
            add
            {
                bool flag = this.AllowNew;
                _onAddingNew = (AddingNewEventHandler)Delegate.Combine(_onAddingNew, value);
                if (flag != this.AllowNew)
                {
                    this.FireListChanged(ListChangedType.Reset, -1);
                }
            }
            remove
            {
                bool flag = this.AllowNew;
                _onAddingNew = (AddingNewEventHandler)Delegate.Remove(_onAddingNew, value);
                if (flag != this.AllowNew)
                    this.FireListChanged(ListChangedType.Reset, -1);
            }
        }

        /// <summary>Tritt ein, wenn die Liste oder ein Element der Liste geändert wird.</summary>
        public event ListChangedEventHandler ListChanged
        {
            add
            {
                _onListChanged = (ListChangedEventHandler)Delegate.Combine(_onListChanged, value);
            }
            remove
            {
                _onListChanged = (ListChangedEventHandler)Delegate.Remove(_onListChanged, value);
            }
        }

        #endregion

        #region Public Properties

        /// <summary>Ruft den Wert ab, der angibt, ob durch das Hinzufügen oder Entfernen von Elementen in der Liste <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignisse ausgelöst werden, oder legt diesen Wert fest.</summary>
        /// <returns>true, wenn durch das Hinzufügen oder Löschen von Elementen <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignisse ausgelöst werden, andernfalls false. Der Standardwert ist true.</returns>
        public bool RaiseListChangedEvents
        {
            get
            {
                return _raiseListChangedEvents;
            }
            set
            {
                if (_raiseListChangedEvents != value)
                    _raiseListChangedEvents = value;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob der Liste mithilfe der <see cref="M:System.ComponentModel.BindingList`1.AddNew" />-Methode neue Elemente hinzugefügt werden können.</summary>
        /// <returns>true, wenn der Liste mithilfe der <see cref="M:System.ComponentModel.BindingList`1.AddNew" />-Methode neue Elemente hinzugefügt werden können, andernfalls false. Der Standardwert hängt von dem in der Liste enthaltenen zugrunde liegenden Typ ab.</returns>
        public bool AllowNew
        {
            get
            {
                if (_userSetAllowNew || _allowNew)
                    return _allowNew;
                return this.AddingNewHandled;
            }
            set
            {
                bool flag = _allowNew;
                _userSetAllowNew = true;
                _allowNew = value;
                if (flag != value)
                    this.FireListChanged(ListChangedType.Reset, -1);
            }
        }
      
        /// <summary>Ruft einen Wert ab, der angibt, ob Elemente in der Liste bearbeitet werden können, oder legt diesen fest.</summary>
        /// <returns>true, wenn die Listenelemente bearbeitet werden können, andernfalls false. Der Standardwert ist true.</returns>
        public bool AllowEdit
        {
            get
            {
                return _allowEdit;
            }
            set
            {
                if (_allowEdit != value)
                {
                    _allowEdit = value;
                    this.FireListChanged(ListChangedType.Reset, -1);
                }
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob Elemente aus der Auflistung entfernt werden können, oder legt diesen fest. </summary>
        /// <returns>true, wenn Elemente mithilfe der <see cref="M:System.ComponentModel.BindingList`1.RemoveItem(System.Int32)" />-Methode aus der Liste entfernt werden können, andernfalls false. Der Standardwert ist true.</returns>
        public bool AllowRemove
        {
            get
            {
                return _allowRemove;
            }
            set
            {
                if (_allowRemove != value)
                {
                    _allowRemove = value;
                    FireListChanged(ListChangedType.Reset, -1);
                }
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignisse aktiviert sind.</summary>
        /// <returns>true, wenn <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignisse unterstützt werden, andernfalls false. Der Standardwert ist true.</returns>
        public virtual bool SupportsChangeNotificationCore // geändert zu public
        {
            get
            {
                return true;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob die Liste Suchvorgänge unterstützt.</summary>
        /// <returns>true, wenn Liste Suchvorgänge unterstützt, andernfalls false. Der Standardwert ist false.</returns>
        public virtual bool SupportsSearchingCore // geändert zu public
        {
            get
            {
                return false;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob die Liste Sortiervorgänge unterstützt.</summary>
        /// <returns>true, wenn die Liste die Sortierung unterstützt, andernfalls false. Der Standardwert ist false.</returns>
        public virtual bool SupportsSortingCore // geändert zu public 
        {
            get
            {
                return false;
            }
        }

        #endregion

        #region Private Properties

        private bool ItemTypeHasDefaultConstructor
        {
            get
            {
                Type typeFromHandle = typeof(T);
                return typeFromHandle.IsPrimitive ||
                    typeFromHandle.GetConstructor(
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.CreateInstance, null, new Type[0], null) != null;
            }
        }

        private bool AddingNewHandled
        {
            get
            {
                return _onAddingNew != null && _onAddingNew.GetInvocationList().Length > 0;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob die Liste sortiert ist. </summary>
        /// <returns>true, wenn die Liste sortiert ist, andernfalls false. Der Standardwert ist false.</returns>
        public virtual bool IsSortedCore // geändert zu public
        {
            get
            {
                return false;
            }
        }


        /// <summary>Ruft den Eigenschaftendeskriptor auf, mit dem die Liste sortiert wird, wenn die Liste in einer abgeleiteten Klasse sortiert wird, andernfalls null. </summary>
        /// <returns>Der zum Sortieren der Liste verwendete <see cref="T:System.ComponentModel.PropertyDescriptor" />.</returns>
        public virtual PropertyDescriptor SortPropertyCore // geändert zu public
        {
            get
            {
                return null;
            }
        }

        /// <summary>Ruft die Sortierrichtung der Liste ab.</summary>
        /// <returns>Einer der <see cref="T:System.ComponentModel.ListSortDirection" />-Werte. Der Standardwert ist <see cref="F:System.ComponentModel.ListSortDirection.Ascending" />. </returns>
        public virtual ListSortDirection SortDirectionCore // geändert zu public
        {
            get
            {
                return ListSortDirection.Ascending;
            }
        }

        #endregion

        #region IBindingList
        
        /// <summary>Ruft einen Wert ab, der angibt, ob der Liste mithilfe der <see cref="M:System.ComponentModel.BindingList`1.AddNew" />-Methode neue Elemente hinzugefügt werden können.</summary>
        /// <returns>true, wenn der Liste mithilfe der <see cref="M:System.ComponentModel.BindingList`1.AddNew" />-Methode neue Elemente hinzugefügt werden können, andernfalls false. Der Standardwert hängt von dem in der Liste enthaltenen zugrunde liegenden Typ ab.</returns>
        bool IBindingList.AllowNew
        {
            get
            {
                return AllowNew;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob Elemente in der Liste bearbeitet werden können.</summary>
        /// <returns>true, wenn die Listenelemente bearbeitet werden können, andernfalls false. Der Standardwert ist true.</returns>
        bool IBindingList.AllowEdit
        {
            get
            {
                return this.AllowEdit;
            }
        }

        /// <summary>Ruft einen Wert ab, der angibt, ob Elemente aus der Liste entfernt werden können.</summary>
        /// <returns>true, wenn Elemente mithilfe der <see cref="M:System.ComponentModel.BindingList`1.RemoveItem(System.Int32)" />-Methode aus der Liste entfernt werden können, andernfalls false. Der Standardwert ist true.</returns>
        bool IBindingList.AllowRemove
        {
            get
            {
                return AllowRemove;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.SupportsChangeNotification" />.</summary>
        /// <returns>true, wenn bei Änderungen der Liste oder eines Elements ein <see cref="E:System.ComponentModel.IBindingList.ListChanged" />-Ereignis ausgelöst wird, andernfalls false.</returns>
        bool IBindingList.SupportsChangeNotification
        {
            get
            {
                return SupportsChangeNotificationCore;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.SupportsSearching" />.</summary>
        /// <returns>true, wenn die Liste die Suche mit der <see cref="M:System.ComponentModel.IBindingList.Find(System.ComponentModel.PropertyDescriptor,System.Object)" />-Methode unterstützt, andernfalls false.</returns>
        bool IBindingList.SupportsSearching
        {
            get
            {
                return SupportsSearchingCore;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.SupportsSorting" />.</summary>
        /// <returns>true, wenn die Liste die Sortierung unterstützt, andernfalls false.</returns>
        bool IBindingList.SupportsSorting
        {
            get
            {
                return SupportsSortingCore;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.IsSorted" />.</summary>
        /// <returns>true, wenn <see cref="M:System.ComponentModel.IBindingListView.ApplySort(System.ComponentModel.ListSortDescriptionCollection)" /> aufgerufen wurde und <see cref="M:System.ComponentModel.IBindingList.RemoveSort" /> nicht aufgerufen wurde, andernfalls false.</returns>
        bool IBindingList.IsSorted
        {
            get
            {
                return IsSortedCore;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.SortProperty" />.</summary>
        /// <returns>Der <see cref="T:System.ComponentModel.PropertyDescriptor" />, der für die Sortierung verwendet wird.</returns>
        PropertyDescriptor IBindingList.SortProperty
        {
            get
            {
                return SortPropertyCore;
            }
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="P:System.ComponentModel.IBindingList.SortDirection" />.</summary>
        /// <returns>Einer der <see cref="T:System.ComponentModel.ListSortDirection" />-Werte.</returns>
        ListSortDirection IBindingList.SortDirection
        {
            get
            {
                return SortDirectionCore;
            }
        }

        /// <summary>Fügt der Liste ein neues Element hinzu. Weitere Informationen finden Sie unter <see cref="M:System.ComponentModel.IBindingList.AddNew" />.</summary>
        /// <returns>Das der Liste hinzugefügte Element.</returns>
        /// <exception cref="T:System.NotSupportedException">Diese Methode wird nicht unterstützt. </exception>
        object IBindingList.AddNew()
        {
            object obj = this.AddNewCore();
            _addNewPos = ((obj != null) ? base.IndexOf((T)((object)obj)) : -1);
            return obj;
        }

        /// <summary>Sortiert die Liste entsprechend einem <see cref="T:System.ComponentModel.PropertyDescriptor" /> und einer <see cref="T:System.ComponentModel.ListSortDirection" />. Eine ausführliche Beschreibung dieses Members finden Sie unter <see cref="M:System.ComponentModel.IBindingList.ApplySort(System.ComponentModel.PropertyDescriptor,System.ComponentModel.ListSortDirection)" />.</summary>
        /// <param name="prop">Der <see cref="T:System.ComponentModel.PropertyDescriptor" />, nach dem sortiert werden soll.</param>
        /// <param name="direction">Einer der <see cref="T:System.ComponentModel.ListSortDirection" />-Werte.</param>
        void IBindingList.ApplySort(PropertyDescriptor prop, ListSortDirection direction)
        {
            ApplySortCore(prop, direction);
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter der <see cref="M:System.ComponentModel.IBindingList.RemoveSort" />-Methode.</summary>
        void IBindingList.RemoveSort()
        {
            RemoveSortCore();
        }
        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="M:System.ComponentModel.IBindingList.Find(System.ComponentModel.PropertyDescriptor,System.Object)" />.</summary>
        /// <returns>Der Index der Zeile mit dem angegebenen <see cref="T:System.ComponentModel.PropertyDescriptor" />.</returns>
        /// <param name="prop">Der <see cref="T:System.ComponentModel.PropertyDescriptor" />, in dem gesucht werden soll.</param>
        /// <param name="key">Der Wert des property-Parameters, nach dem gesucht werden soll.</param>
        int IBindingList.Find(PropertyDescriptor prop, object key)
        {
            return FindCore(prop, key);
        }

        /// <summary>Sucht nach dem Index des Elements, das über den angegebenen Eigenschaftendeskriptor mit dem angegebenen Wert verfügt, wenn der Suchvorgang in einer abgeleiteten Klasse implementiert wird, andernfalls <see cref="T:System.NotSupportedException" />.</summary>
        /// <returns>Der nullbasierte Index des Elements, das dem Eigenschaftendeskriptor entspricht und den angegebenen Wert enthält.</returns>
        /// <param name="prop">Der zu suchende <see cref="T:System.ComponentModel.PropertyDescriptor" />.</param>
        /// <param name="key">Der Wert von <paramref name="property" />, der übereinstimmen soll.</param>
        /// <exception cref="T:System.NotSupportedException">
        ///   <see cref="M:System.ComponentModel.BindingList`1.FindCore(System.ComponentModel.PropertyDescriptor,System.Object)" /> wird in einer abgeleiteten Klasse nicht überschrieben.</exception>
        public virtual int FindCore(PropertyDescriptor prop, object key) // geändert zu public
        {
            throw new NotSupportedException();
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="M:System.ComponentModel.IBindingList.AddIndex(System.ComponentModel.PropertyDescriptor)" />.</summary>
        /// <param name="prop">Der als Suchkriterium hinzuzufügende <see cref="T:System.ComponentModel.PropertyDescriptor" />. </param>
        void IBindingList.AddIndex(PropertyDescriptor prop)
        {
        }

        /// <summary>Eine Beschreibung dieses Members finden Sie unter <see cref="M:System.ComponentModel.IBindingList.RemoveIndex(System.ComponentModel.PropertyDescriptor)" />.</summary>
        /// <param name="prop">Ein <see cref="T:System.ComponentModel.PropertyDescriptor" />, der aus den für die Suche verwendeten Indizes entfernt werden soll.</param>
        void IBindingList.RemoveIndex(PropertyDescriptor prop)
        {
        }

        #endregion

        #region IRaiseItemChangedEvents

        /// <summary>Ruft einen Wert ab, der angibt, ob durch Änderungen des Elementeigenschaftenwerts <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignisse vom Typ <see cref="F:System.ComponentModel.ListChangedType.ItemChanged" /> ausgelöst werden. Dieser Member kann in einer abgeleiteten Klasse nicht überschrieben werden.</summary>
        /// <returns>true , wenn der Listentyp <see cref="T:System.ComponentModel.INotifyPropertyChanged" /> implementiert, andernfalls false. Der Standardwert ist false.</returns>
        bool IRaiseItemChangedEvents.RaisesItemChangedEvents
        {
            get
            {
                return _raiseItemChangedEvents;
            }
        }

        #endregion

        #region Virtual Methods

        /// <summary>Löst das <see cref="E:System.ComponentModel.BindingList`1.AddingNew" />-Ereignis aus.</summary>
        /// <param name="e">Eine Instanz von <see cref="T:System.ComponentModel.AddingNewEventArgs" />, die die Ereignisdaten enthält. </param>
        public virtual void OnAddingNew(AddingNewEventArgs e) // geändert zu public
        {
            if (_onAddingNew != null)
                _onAddingNew(this, e);
        }
        /// <summary>Löst das <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignis aus.</summary>
        /// <param name="e">Eine Instanz von <see cref="T:System.ComponentModel.ListChangedEventArgs" />, die die Ereignisdaten enthält. </param>
        public virtual void OnListChanged(ListChangedEventArgs e) // geändert zu public
        {
            if (_onListChanged != null)
                _onListChanged(this, e);
        }

        /// <summary>Entfernt alle Elemente aus der Auflistung.</summary>
        protected override void ClearItems()
        {
            this.EndNew(_addNewPos);
            if (_raiseItemChangedEvents)
            {
                foreach (T current in base.Items)
                    this.UnhookPropertyChanged(current);
            }
            base.ClearItems();
            this.FireListChanged(ListChangedType.Reset, -1);
        }

        /// <summary>Fügt das angegebene Element am angegebenen Index in die Liste ein.</summary>
        /// <param name="index">Der nullbasierte Index, an dem das Element eingefügt werden soll.</param>
        /// <param name="item">Das in die Liste einzufügende Element.</param>
        protected override void InsertItem(int index, T item)
        {
            this.EndNew(_addNewPos);
            base.InsertItem(index, item);
            if (_raiseItemChangedEvents)
                this.HookPropertyChanged(item);
            this.FireListChanged(ListChangedType.ItemAdded, index);
        }

        /// <summary>Entfernt das Element am angegebenen Index.</summary>
        /// <param name="index">Der nullbasierte Index des zu entfernenden Elements. </param>
        /// <exception cref="T:System.NotSupportedException">Sie entfernen ein neu hinzugefügtes Element, und <see cref="P:System.ComponentModel.IBindingList.AllowRemove" /> ist auf false festgelegt. </exception>
        protected override void RemoveItem(int index)
        {
            if (!_allowRemove && (_addNewPos < 0 || _addNewPos != index))
                throw new NotSupportedException();

            this.EndNew(_addNewPos);
            if (_raiseItemChangedEvents)
                this.UnhookPropertyChanged(base[index]);

            base.RemoveItem(index);
            this.FireListChanged(ListChangedType.ItemDeleted, index);
        }

        /// <summary>Verwirft ein ausstehendes neues Element.</summary>
        /// <param name="itemIndex">Der Index des neuen hinzuzufügenden Elements. </param>
        public virtual void CancelNew(int itemIndex)
        {
            if (_addNewPos >= 0 && _addNewPos == itemIndex)
            {
                this.RemoveItem(_addNewPos);
                _addNewPos = -1;
            }
        }

        /// <summary>Ersetzt das Element an der angegebenen Position durch ein angegebenes Element.</summary>
        /// <param name="index">Der nullbasierte Index des zu ersetzenden Elements.</param>
        /// <param name="item">Der neue Wert für das Element am angegebenen Index. Der Wert kann für Referenztypen null sein.</param>
        /// <exception cref="T:System.ArgumentOutOfRangeException">
        ///   <paramref name="index" /> ist kleiner als 0 (null).– oder –<paramref name="index" /> ist größer als <see cref="P:System.Collections.ObjectModel.Collection`1.Count" />.</exception>
        protected override void SetItem(int index, T item) // geändert zu public
        {
            if (_raiseItemChangedEvents)
                this.UnhookPropertyChanged(base[index]);

            base.SetItem(index, item);

            if (_raiseItemChangedEvents)
                this.HookPropertyChanged(item);

            this.FireListChanged(ListChangedType.ItemChanged, index);
        }

        /// <summary>Führt einen Commit eines ausstehenden neuen Elements für die Auflistung aus.</summary>
        /// <param name="itemIndex">Der Index des neuen hinzuzufügenden Elements.</param>
        public virtual void EndNew(int itemIndex)
        {
            if (_addNewPos >= 0 && _addNewPos == itemIndex)
                _addNewPos = -1;
        }

        /// <summary>Fügt am Ende der Auflistung ein neues Element hinzu.</summary>
        /// <returns>Das der Auflistung hinzugefügte Element.</returns>
        /// <exception cref="T:System.InvalidCastException">Der Typ des neuen Elements entspricht nicht dem Typ der in <see cref="T:System.ComponentModel.BindingList`1" /> enthaltenen Objekte.</exception>
        public virtual object AddNewCore() // geändert zu public
        {
            object obj = this.FireAddingNew();
            if (obj == null)
            {
                Type typeFromHandle = typeof(T);
                obj = SecureCreateInstance(typeFromHandle);
            }
            base.Add((T)((object)obj));
            return obj;
        }

        /// <summary>Sortiert die gegebenenfalls in einer abgeleiteten Klasse überschriebenen Elemente; andernfalls wird eine <see cref="T:System.NotSupportedException" /> ausgelöst.</summary>
        /// <param name="prop">Ein <see cref="T:System.ComponentModel.PropertyDescriptor" />, der die Eigenschaft angibt, nach der sortiert werden soll.</param>
        /// <param name="direction">Einer der <see cref="T:System.ComponentModel.ListSortDirection" />-Werte.</param>
        /// <exception cref="T:System.NotSupportedException">Die Methode wird in einer abgeleiteten Klasse nicht überschrieben. </exception>
        public virtual void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction) // geändert zu public
        {
            throw new NotSupportedException();
        }

        /// <summary>Entfernt jede mit <see cref="M:System.ComponentModel.BindingList`1.ApplySortCore(System.ComponentModel.PropertyDescriptor,System.ComponentModel.ListSortDirection)" /> angewendete Sortierung, wenn die Sortierung in einer abgeleiteten Klasse implementiert wird; andernfalls wird <see cref="T:System.NotSupportedException" /> ausgelöst.</summary>
        /// <exception cref="T:System.NotSupportedException">Die Methode wird in einer abgeleiteten Klasse nicht überschrieben. </exception>
        public virtual void RemoveSortCore() // geändert zu public
        {
            throw new NotSupportedException();
        }

        #endregion

        #region Public Methods

        /// <summary>Löst für das Element an der angegebenen Position ein <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignis vom Typ <see cref="F:System.ComponentModel.ListChangedType.ItemChanged" /> aus.</summary>
        /// <param name="position">Ein nullbasierter Index des zurückzusetzenden Elements.</param>
        public void ResetItem(int position)
        {
            FireListChanged(ListChangedType.ItemChanged, position);
        }

        public void CallClearItems()
        {
            ClearItems();
        }

        public void CallInsertItem(int index, T item) // geändert zu public
        {
            InsertItem(index, item);
        }

        public void CallRemoveItem(int index)
        {
            RemoveItem(index);
        }

        public void CallSetItem(int index, T item)
        {
            SetItem(index, item);
        }

        /// <summary>Fügt der Auflistung ein neues Element hinzu.</summary>
        /// <returns>Das der Liste hinzugefügte Element.</returns>
        /// <exception cref="T:System.InvalidOperationException">Die <see cref="P:System.Windows.Forms.BindingSource.AllowNew" />-Eigenschaft ist auf false festgelegt. – oder –Für den aktuellen Elementtyp konnte kein öffentlicher Standardkonstruktor gefunden werden.</exception>
        public T AddNew()
        {
            return (T)((object)((IBindingList)this).AddNew());
        }

        #endregion

        #region Private Methods
        
        private object FireAddingNew()
        {
            AddingNewEventArgs addingNewEventArgs = new AddingNewEventArgs(null);
            this.OnAddingNew(addingNewEventArgs);
            return addingNewEventArgs.NewObject;
        }

        /// <summary>Löst ein <see cref="E:System.ComponentModel.BindingList`1.ListChanged" />-Ereignis vom Typ <see cref="F:System.ComponentModel.ListChangedType.Reset" /> aus.</summary>
        public void ResetBindings()
        {
            FireListChanged(ListChangedType.Reset, -1);
        }

        public void RaiseListChanged(ListChangedType type, int index)
        {
            if (_raiseListChangedEvents)
                this.OnListChanged(new ListChangedEventArgs(type, index));
        }

        protected internal override void FireListChanged(ListChangedType type, int index)
        {
            if (_raiseListChangedEvents)
                this.OnListChanged(new ListChangedEventArgs(type, index));
        }
      
        private void HookPropertyChanged(T item)
        {
            INotifyPropertyChanged notifyPropertyChanged = item as INotifyPropertyChanged;
            if (notifyPropertyChanged != null)
            {
                if (_propertyChangedEventHandler == null)
                {
                    _propertyChangedEventHandler = new PropertyChangedEventHandler(this.Child_PropertyChanged);
                }
                notifyPropertyChanged.PropertyChanged += _propertyChangedEventHandler;
            }
        }

        private void UnhookPropertyChanged(T item)
        {
            INotifyPropertyChanged notifyPropertyChanged = item as INotifyPropertyChanged;
            if (notifyPropertyChanged != null && _propertyChangedEventHandler != null)
            {
                notifyPropertyChanged.PropertyChanged -= _propertyChangedEventHandler;
            }
        }

        private void Child_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (this.RaiseListChangedEvents)
            {
                if (sender == null || e == null || string.IsNullOrEmpty(e.PropertyName))
                {
                    this.ResetBindings();
                    return;
                }
                T t;
                try
                {
                    t = (T)((object)sender);
                }
                catch (InvalidCastException)
                {
                    this.ResetBindings();
                    return;
                }
                int num = _lastChangeIndex;
                if (num >= 0 && num < base.Count)
                {
                    T t2 = base[num];
                    if (t2.Equals(t))
                    {
                        goto IL_7B;
                    }
                }
                num = base.IndexOf(t);
                _lastChangeIndex = num;
            IL_7B:
                if (num == -1)
                {
                    this.UnhookPropertyChanged(t);
                    this.ResetBindings();
                    return;
                }
                if (_itemTypeProperties == null)
                {
                    _itemTypeProperties = TypeDescriptor.GetProperties(typeof(T));
                }
                PropertyDescriptor propDesc = _itemTypeProperties.Find(e.PropertyName, true);
                ListChangedEventArgs e2 = new ListChangedEventArgs(ListChangedType.ItemChanged, num, propDesc);
                this.OnListChanged(e2);
            }
        }

        #endregion

        #region SecurtityUtils

        internal object SecureCreateInstance(Type type)
        {
            return SecureCreateInstance(type, null);
        }

        internal object SecureCreateInstance(Type type, object[] args)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            ResolveArgumentsOnCreateNew(ref args);

            if (type.Assembly == typeof(BindingList<T>).Assembly && !type.IsPublic && !type.IsNestedPublic)
                new ReflectionPermission(PermissionState.Unrestricted).Demand();
            
            return Activator.CreateInstance(type, args);
        }
 
        #endregion

        /// <summary>
        /// Resolve custom ctor item arguments for the IBindingList.AddNew method
        /// </summary>
        /// <param name="args">arguments array for the item</param>
        protected internal virtual void ResolveArgumentsOnCreateNew(ref object[] args)
        {

        }
    }
}
