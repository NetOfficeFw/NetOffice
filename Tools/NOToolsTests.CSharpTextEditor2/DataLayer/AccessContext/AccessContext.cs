using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Represents an isolated store/cache to create local changes and commit these changes in a transaction or cancel changes anytime
    /// </summary>
    public class AccessContext : IEnumerable<AccessContextList>, INotifyPropertyChanged
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">parent collection</param>
        /// <param name="name">unique name of the context</param>
        internal AccessContext(AccessContextCollection parent, string name)
        {
            Name = name;
            Parent = parent;
            Commands = new Chain<Command>();
            Tables = new AccessContextListCollection(this);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Unique name of the context
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Get a proxy for a root table
        /// </summary>
        /// <param name="tableName">name of the root table</param>
        /// <returns>root table instance</returns>
        public AccessContextList this[string tableName]
        {
            get 
            {
                foreach (AccessContextList item in Tables)
                {
                    if (item.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }

                foreach (RootList item in Parent.DataSources)
                {
                    if (item.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AccessContextList table = new AccessContextList(this, item);

                        table.AfterAddInsert += new NOTools.ComponentModel.Collection<AccessContextItem>.AfterAddInsertEventHandler(Table_AfterAddInsert);
                        table.AfterRemove += new NOTools.ComponentModel.Collection<AccessContextItem>.AfterRemoveEventHandler(Table_AfterRemove);
                       
                        Tables.Add(table);
                        return table;                        
                    }
                }

                throw new ArgumentOutOfRangeException(tableName);
            }
        }

        /// <summary>
        /// Returns info the context contains one or more item(s) with local changes
        /// </summary>
        public bool ContainsLocalChanges
        {
            get
            {
                foreach (AccessContextList item in Tables)
                    if (item.ContainsLocalChanges)
                        return true;
                return false;
            }
        }

        /// <summary>
        /// Command collection to provide undo/redo
        /// </summary>
        public Chain<Command> Commands { get; private set; }

        /// <summary>
        /// Parent collection
        /// </summary>
        internal AccessContextCollection Parent { get; private set; }

        /// <summary>
        /// Current created root table proxies
        /// </summary>
        internal AccessContextListCollection Tables { get; private set; }

        /// <summary>
        /// Returns info the access context is currently proceed a undo/redo action
        /// </summary>
        internal bool IsCurrentlyInUndoRedoAction { get; set; }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Occures when the ContainsLocalChanges property has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        internal void RaiseNotifyPropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Methods

        /// <summary>
        /// Reload all associated root tables from database and reset local proxies
        /// </summary>
        public void UpdateDataSourcesAndResetLocalData()
        {
            // to do: implement this
            throw new NotSupportedException();
        }

        /// <summary>
        /// Reset all local proxies an reload data from root tables
        /// </summary>
        public void ResetLocalData()        
        {
            foreach (AccessContextList item in Tables)
                item.ResetLocalData();
            Commands.Clear();
        }

        /// <summary>
        /// Commit all local changes to root tables
        /// </summary>
        public void ApplyLocalChanges()
        {
            foreach (AccessContextList item in this)
                item.ApplyLocalChanges();
            Commands.Clear();
        }

        /// <summary>
        /// Rollback all local changes. The method doesnt reload data from root tables
        /// </summary>
        public void CancelLocalChanges()
        {
            foreach (AccessContextList item in this)
                item.CancelLocalChanges();
            Commands.Clear();
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// AccessContextList Enumerator
        /// </summary>
        /// <returns>Enumerator Instance</returns>
        public IEnumerator<AccessContextList> GetEnumerator()
        {
            return Tables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return Tables.GetEnumerator();
        }

        #endregion
        
        #region Trigger

        internal void Item_IsLocalChanged(AccessContextItem item, string propertyName, object oldValue, object newValue)
        {
            if (IsCurrentlyInUndoRedoAction)
                return;

            var changeRedo = ChangeClosure.Create(this, item.SetValue, propertyName, newValue);
            var changeUndo = ChangeClosure.Create(this, item.SetValue, propertyName, oldValue);
            Commands.Append(new Command("Change Item", changeRedo, changeUndo));
        }

        private void Table_AfterAddInsert(AccessContextItem item, int itemIndex)
        {
            if (IsCurrentlyInUndoRedoAction)
                return;

            var addRedo = InsertClosure.Create(this, item.Parent.Insert, itemIndex, item);
            var addUndo = RemoveClosure.Create(this, item.Parent.Remove, item);
            Commands.Append(new Command("Add Item", addRedo, addUndo));
        }

        private void Table_AfterRemove(AccessContextItem item, int itemIndex)
        {
            if (IsCurrentlyInUndoRedoAction)
                return;

            var deleteRedo = RemoveClosure.Create(this, item.Parent.Remove, item);
            var deleteUndo = InsertClosure.Create(this, item.Parent.Insert, itemIndex, item);
            Commands.Append(new Command("Delete Item", deleteRedo, deleteUndo));
        }

        private void Table_ListChanged(object sender, ListChangedEventArgs e)
        {
            switch (e.ListChangedType)
            {
                case ListChangedType.Reset:
                    Commands.Clear();
                    break;
            }
        }

        #endregion
    }
}
