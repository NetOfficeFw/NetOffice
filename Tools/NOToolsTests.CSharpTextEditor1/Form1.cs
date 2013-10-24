using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using NOTools.InMemoryCompiler;

namespace NOToolsTests.CSharpTextEditor1
{
    public partial class Form1 : Form
    {
        #region Ctor

        public Form1()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        private CompileResult CompileCode(string[] references)
        {
            DynamicAssembly assembly = new DynamicAssembly("AnyAssemblyName", references);
            assembly.CustomClasses.AddNew(codeEditorControl1.Text);
            CompileResult result = CSharpCompiler.CompileDynamicAssembly(assembly);
            return result;
        }

        private void RunCode(Assembly assembly)
        {
            try
            {
                Type[] types = assembly.GetExportedTypes();
                foreach (Type type in types)
                {
                    Type interfaceType = type.GetInterface("NOToolsTests.CSharpTextEditor1.IScriptComponent");
                    if (null != interfaceType)
                    {
                        IScriptComponent component = assembly.CreateInstance(type.FullName) as IScriptComponent;
                        component.Execute();
                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                System.CodeDom.Compiler.CompilerErrorCollection collection = new System.CodeDom.Compiler.CompilerErrorCollection();
                System.CodeDom.Compiler.CompilerError error = new System.CodeDom.Compiler.CompilerError("", 0, 0, "0", exception.Message);
                collection.Add(error);
                codeEditorControl1.ShowErrors(collection);
            }
        }

        #endregion

        #region Trigger

        private void StripButtonOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter = "C# Files(*.cs)|*.cs";
            if (DialogResult.OK == dialog.ShowDialog(this))
                codeEditorControl1.Text = System.IO.File.ReadAllText(dialog.FileName, Encoding.UTF8);
        }

        private void StripButtonSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "C# Files(*.cs)|*.cs";
            if (DialogResult.OK == dialog.ShowDialog(this))
            {
                if (System.IO.File.Exists(dialog.FileName))
                    System.IO.File.Delete(dialog.FileName);
                 System.IO.File.AppendAllText(dialog.FileName, codeEditorControl1.Text, Encoding.UTF8);
            }
        }

        private void StripButtonNew_Click(object sender, EventArgs e)
        {
            codeEditorControl1.Text = String.Empty;
        }

        private void StripButtonRun_Click(object sender, EventArgs e)
        {
            codeEditorControl1.ErrorPanelSettings.Header = "Compile...";
            CompileResult result = CompileCode(codeEditorControl1.References.ToStringPathArray());
            codeEditorControl1.ShowErrors(result.Errors, "Sucseed");
            if (result.Errors.Count == 0 && null != result.Assembly)
                RunCode(result.Assembly);
        }

        private void StripButtonCompile_Click(object sender, EventArgs e)
        {
            codeEditorControl1.ErrorPanelSettings.Header = "Compile...";
            CompileResult result = CompileCode(codeEditorControl1.References.ToStringPathArray());
            codeEditorControl1.ShowErrors(result.Errors, "Sucseed");
        }

        private void StripButtonAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "CSharpTextEditor Test Application", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void codeEditorControl1_CompileRequest(NOTools.CSharpTextEditor.CodeEditorControl sender, NOTools.CSharpTextEditor.CompileRequestEventArgs args)
        { 
            codeEditorControl1.ErrorPanelSettings.Header = "Compile...";
            CompileResult result = CompileCode(codeEditorControl1.References.ToStringPathArray());
            codeEditorControl1.ShowErrors(result.Errors, "Sucseed");
        }

        private void codeEditorControl1_RunRequest(NOTools.CSharpTextEditor.CodeEditorControl sender, NOTools.CSharpTextEditor.CompileRequestEventArgs args)
        {
            codeEditorControl1.ErrorPanelSettings.Header = "Compile...";
            CompileResult result = CompileCode(codeEditorControl1.References.ToStringPathArray());
            codeEditorControl1.ShowErrors(result.Errors, "Sucseed");
            if (result.Errors.Count == 0 && null != result.Assembly)
                RunCode(result.Assembly);
        }

        private void codeEditorControl1_PersistanceResolve(string name, ref string path)
        {
            if (name == "NOToolsTests.CSharpTextEditor1")
                path = System.IO.Path.Combine(Application.StartupPath, "NOToolsTests.CSharpTextEditor1.exe");
        }

        #endregion
    }
}
