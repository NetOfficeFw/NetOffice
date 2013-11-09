using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    #region Delegates

    /// <summary>
    /// AddNew Handler
    /// </summary>
    /// <returns>new created item</returns>
    public delegate AccessContextItem AddNewHandler();

    /// <summary>
    /// Insert Handler
    /// </summary>
    /// <param name="index">index to insert</param>
    /// <param name="item">target item</param>
    public delegate void InsertHandler(int index, AccessContextItem item);

    /// <summary>
    /// Remove Handler
    /// </summary>
    /// <param name="item">item to remove</param>
    /// <returns>true if removed otherwise false</returns>
    public delegate bool RemoveHandler(AccessContextItem item);
  
    /// <summary>
    /// Set Property Value handler
    /// </summary>
    /// <param name="propertyName">name of the property</param>
    /// <param name="propertyValue">new value of the property</param>
    public delegate void SetValueHandler(string propertyName, object propertyValue);

    #endregion

    /// <summary>
    /// Base class for definded closures
    /// </summary>
    public abstract class Closure
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parentContext">associated context</param>
        public Closure(AccessContext parentContext)
        {
            ParentContext = parentContext;
        }

        /// <summary>
        /// Associated context
        /// </summary>
        public AccessContext ParentContext { get; private set; }

        /// <summary>
        /// Execute the command
        /// </summary>
        public abstract void Execute();
    }

    /// <summary>
    /// AddNew Closure
    /// </summary>
    public class AddNewClosure : Closure
    {
        private AddNewHandler _execute;

        public AddNewClosure(AccessContext parentContext)  : base(parentContext)
        { }

        public static Closure Create(AccessContext parentContext, AddNewHandler action)
        {
            return new AddNewClosure(parentContext) { _execute = () => action() };
        }

        public override void Execute()
        {
            ParentContext.IsCurrentlyInUndoRedoAction = true;
            try
            {
                _execute();
            }
            catch
            {
                throw;
            }
            finally
            {
                ParentContext.IsCurrentlyInUndoRedoAction = false;
            }
        }
    }

    /// <summary>
    /// Insert Closure
    /// </summary>
    public class InsertClosure : Closure
    {
        private InsertHandler _execute;
        private int _arg0;
        private AccessContextItem _arg1;

        public InsertClosure(AccessContext parentContext, InsertHandler action, int index, AccessContextItem item) : base(parentContext)
        {
            _execute = action;
            _arg0 = index;
            _arg1 = item;
        }

        public static Closure Create(AccessContext parentContext, InsertHandler action, int index, AccessContextItem item)
        {
            return new InsertClosure(parentContext, action, index, item);
        }

        public override void Execute()
        {
            ParentContext.IsCurrentlyInUndoRedoAction = true;
            try
            {
                _execute(_arg0, _arg1);
            }
            catch
            {
                throw;
            }
            finally
            {
                ParentContext.IsCurrentlyInUndoRedoAction = false;
            }
        }
    }

    /// <summary>
    /// Remove Closure
    /// </summary>
    public class RemoveClosure : Closure
    {
        private RemoveHandler _execute;
        private AccessContextItem _arg0;

        internal RemoveClosure(AccessContext parentContext, RemoveHandler action, AccessContextItem item) : base(parentContext)
        {
            _execute = action;
            _arg0 = item;
        }

        public static Closure Create(AccessContext parentContext, RemoveHandler action, AccessContextItem item)
        {
            return new RemoveClosure(parentContext, action, item);
        }

        public override void Execute()
        {
            ParentContext.IsCurrentlyInUndoRedoAction = true;
            try
            {
                _execute(_arg0);
            }
            catch 
            {
                throw;
            }
            finally
            {
                ParentContext.IsCurrentlyInUndoRedoAction = false;
            }
        }
    }

    /// <summary>
    /// Change Property Value Closure
    /// </summary>
    public class ChangeClosure : Closure
    {
        private SetValueHandler _execute;
        private string _arg0;
        private object _arg1;

        internal ChangeClosure(AccessContext parentContext, SetValueHandler action, string propertyName, object propertyValue) : base(parentContext)
        {
            _execute = action;
            _arg0 = propertyName;
            _arg1 = propertyValue;
        }

        public static Closure Create(AccessContext parentContext, SetValueHandler action,string propertyName, object propertyValue)
        {
            return new ChangeClosure(parentContext, action, propertyName, propertyValue);
        }

        public override void Execute()
        {
            ParentContext.IsCurrentlyInUndoRedoAction = true;
            try
            {
                _execute(_arg0, _arg1);
            }
            catch
            {
                throw;
            }
            finally
            {
                ParentContext.IsCurrentlyInUndoRedoAction = false;
            }
        }
    }   
}
