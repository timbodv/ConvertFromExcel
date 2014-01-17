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
            this.ErrorOccured = false;
        }

        public bool ErrorOccured { get; private set; }

        public ExcelFormatEnum Format { get; private set; }

        public DataSet Data { get; private set; }

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
            catch (Exception)
            {
                objectToReturn.ErrorOccured = true;
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
                    catch (Exception)
                    {
                        objectToReturn.ErrorOccured = true;
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
                    catch (Exception)
                    {
                        objectToReturn.ErrorOccured = true;
                    }
                }
            }

            return objectToReturn;
        }
    }
}
