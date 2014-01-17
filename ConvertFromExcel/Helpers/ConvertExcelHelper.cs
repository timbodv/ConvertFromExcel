namespace Timbodv.ConvertFromExcel
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Excel;

    public class ConvertExcelHelper
    {
        public ConvertExcelHelper()
        {
            this.Data = new DataSet();
        }

        public Exception Error { get; set; }

        public ExcelFormatEnum Format { get; set; }

        public DataSet Data { get; set; }

        public static ConvertExcelHelper ConvertExcelToDataset(string filename, bool includesHeader)
        {
            ConvertExcelHelper objectToReturn = new ConvertExcelHelper();
            bool openedFile = false;

            try
            {
                using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read))
                {
                    if (ExcelReaderFactory.CreateOpenXmlReader(stream).IsValid)
                    {
                        objectToReturn.Format = ExcelFormatEnum.XLSX;
                    }
                    else
                    {
                        objectToReturn.Format = ExcelFormatEnum.XLS;
                    }

                    openedFile = true;
                }
            }
            catch (Exception e)
            {
                objectToReturn.Error = e;
            }

            if (openedFile)
            {
                if (objectToReturn.Format == ExcelFormatEnum.XLSX)
                {
                    try
                    {
                        using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read))
                        {
                            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            excelReader.IsFirstRowAsColumnNames = includesHeader;
                            objectToReturn.Data = excelReader.AsDataSet();
                            excelReader.Close();
                        }
                    }
                    catch (Exception e)
                    {
                        objectToReturn.Error = e;
                    }
                }
                else
                {
                    // must be XLS
                    try
                    {
                        using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read))
                        {
                            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                            excelReader.IsFirstRowAsColumnNames = includesHeader;
                            objectToReturn.Data = excelReader.AsDataSet();
                            excelReader.Close();
                        }
                    }
                    catch (Exception e)
                    {
                        objectToReturn.Error = e;
                    }
                }
            }

            return objectToReturn;
        }
    }
}
