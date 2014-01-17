namespace Timbodv.ConvertFromExcel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Management;
    using System.Management.Automation;
    using System.Text;
    using System.Threading.Tasks;

    [Cmdlet(VerbsData.ConvertFrom, "Excel")]
    public class ConvertFromExcelCommand : PSCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public string FileName { get; set; }

        [Parameter]
        public SwitchParameter IncludesHeader { get; set; }

        protected override void ProcessRecord()
        {
            ConvertExcelHelper convert = ConvertExcelHelper.ConvertExcelToDataset(this.FileName, this.IncludesHeader);
            if (!convert.ErrorOccured)
            {
                this.WriteVerbose(@"Excel document format was " + convert.Format);
                this.WriteObject(convert.Data);
            }
            else
            {
                ErrorRecord record = new ErrorRecord(new ApplicationException("Failed to parse specified file"), "ParserError,Timbodv.ConvertFromExcel.Command.ConvertFromExcelCommand", ErrorCategory.ParserError, this);
                this.WriteError(record);
            }
        }
    }
}
