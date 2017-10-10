using System;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Tools;
using MSForms = NetOffice.MSFormsApi;
using NetOffice.OutlookApi.Tools.Contribution;
using NetOffice.Extensions.Conversion;

namespace Outlook06AddinCS4
{
    public class CustomFormRegion1 : OpenFormRegion
    {
        private Outlook.OlkTextBox _textBox1;
        private Outlook.OlkCommandButton _commandButton1;

        public CustomFormRegion1(Outlook.FormRegion formRegion) : base(formRegion)
        {
            MSForms.UserForm form = formRegion.Form as MSForms.UserForm;
            _textBox1 = form.Controls["TextBox1"].To<Outlook.OlkTextBox>();
            _commandButton1 = form.Controls["CommandButton1"].To<Outlook.OlkCommandButton>();
            
            if(null != _commandButton1)     
                _commandButton1.ClickEvent += CommandButton1_ClickEvent;
        }

        private void CommandButton1_ClickEvent()
        {
            if(null != _textBox1)
                OutlookDialogUtils.ShowMessageBox(_textBox1.Text, "Outlook06AddinCS4");
        }
    }
}
