using System.Collections.Generic;
using System.Dynamic;

namespace NetOffice.Dynamics
{
    /*
        RuntimeBinder may throws some trial/error exceptions while bind to instance.
    */

    /// <summary>
    /// Wrapper arround underylying DynamicMetaObject for debugging purpose
    /// </summary>
    public class COMDynamicMetaObject : DynamicMetaObject
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="underlying">wrapped instance</param>
        public COMDynamicMetaObject(DynamicMetaObject underlying) : base(underlying.Expression, underlying.Restrictions, underlying.Value)
        {
            Underlying = underlying;
        }

        /// <summary>
        /// Wrapped Instance
        /// </summary>
        private DynamicMetaObject Underlying { get; set; }

        /// <summary>
        /// Performs the binding of the dynamic binary operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.BinaryOperationBinder that represents the details of the dynamic operation.</param>
        /// <param name="arg"> An instance of the System.Dynamic.DynamicMetaObject representing the right hand side of the binary operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindBinaryOperation(BinaryOperationBinder binder, DynamicMetaObject arg)
        {
            return Underlying.BindBinaryOperation(binder, arg);
        }

        /// <summary>
        /// Performs the binding of the dynamic conversion operation.
        /// </summary>
        /// <param name="binder"> An instance of the System.Dynamic.ConvertBinder that represents the details of the dynamic operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindConvert(ConvertBinder binder)
        {
            return Underlying.BindConvert(binder);
        }

        /// <summary>
        /// Performs the binding of the dynamic create instance operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.CreateInstanceBinder that represents the details of the dynamic operation.</param>
        /// <param name="args">An array of System.Dynamic.DynamicMetaObject instances - arguments to the create instance operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindCreateInstance(CreateInstanceBinder binder, DynamicMetaObject[] args)
        {
            return Underlying.BindCreateInstance(binder, args);
        }

        /// <summary>
        /// Performs the binding of the dynamic delete index operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.DeleteIndexBinder that represents the details of the dynamic operation.</param>
        /// <param name="indexes">An array of System.Dynamic.DynamicMetaObject instances - indexes for the delete index operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindDeleteIndex(DeleteIndexBinder binder, DynamicMetaObject[] indexes)
        {
            return Underlying.BindDeleteIndex(binder, indexes);
        }

        /// <summary>
        /// Performs the binding of the dynamic delete member operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.DeleteMemberBinder that represents the details of the dynamic operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindDeleteMember(DeleteMemberBinder binder)
        {
            return Underlying.BindDeleteMember(binder);
        }

        /// <summary>
        /// Performs the binding of the dynamic get index operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.GetIndexBinder that represents the details of the dynamic operation.</param>
        /// <param name="indexes">An array of System.Dynamic.DynamicMetaObject instances - indexes for the get index operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindGetIndex(GetIndexBinder binder, DynamicMetaObject[] indexes)
        {
            return Underlying.BindGetIndex(binder, indexes);
        }

        /// <summary>
        /// Performs the binding of the dynamic get member operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.GetMemberBinder that represents the details of the dynamic operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindGetMember(GetMemberBinder binder)
        {
            return Underlying.BindGetMember(binder);
        }

        /// <summary>
        /// Performs the binding of the dynamic invoke operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.InvokeBinder that represents the details of the dynamic operation.</param>
        /// <param name="args">An array of System.Dynamic.DynamicMetaObject instances - arguments to the invoke operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindInvoke(InvokeBinder binder, DynamicMetaObject[] args)
        {
            return Underlying.BindInvoke(binder, args);
        }

        /// <summary>
        /// Performs the binding of the dynamic invoke member operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.InvokeMemberBinder that represents the details of the dynamic operation.</param>
        /// <param name="args">An array of System.Dynamic.DynamicMetaObject instances - arguments to the invoke member operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindInvokeMember(InvokeMemberBinder binder, DynamicMetaObject[] args)
        {
            return Underlying.BindInvokeMember(binder, args);
        }

        /// <summary>
        /// Performs the binding of the dynamic set index operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.SetIndexBinder that represents the details of the dynamic operation.</param>
        /// <param name="indexes">An array of System.Dynamic.DynamicMetaObject instances - indexes for the set index operation.</param>
        /// <param name="value">The System.Dynamic.DynamicMetaObject representing the value for the set index operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindSetIndex(SetIndexBinder binder, DynamicMetaObject[] indexes, DynamicMetaObject value)
        {
            return Underlying.BindSetIndex(binder, indexes, value);
        }

        /// <summary>
        /// Performs the binding of the dynamic set member operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.SetMemberBinder that represents the details of the dynamic operation.</param>
        /// <param name="value">The System.Dynamic.DynamicMetaObject representing the value for the set member  operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindSetMember(SetMemberBinder binder, DynamicMetaObject value)
        {
            return Underlying.BindSetMember(binder, value);
        }

        /// <summary>
        /// Performs the binding of the dynamic unary operation.
        /// </summary>
        /// <param name="binder">An instance of the System.Dynamic.UnaryOperationBinder that represents the details of the dynamic operation.</param>
        /// <returns>The new System.Dynamic.DynamicMetaObject representing the result of the binding.</returns>
        public override DynamicMetaObject BindUnaryOperation(UnaryOperationBinder binder)
        {
            return Underlying.BindUnaryOperation(binder);
        }

        /// <summary>
        /// Returns the enumeration of all dynamic member names.
        /// </summary>
        /// <returns>The list of dynamic member names.</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return Underlying.GetDynamicMemberNames();
        }
    }
}