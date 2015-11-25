using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Reflection;
using System.IO;

using Microsoft.Office.Interop.Excel;


public partial class MicrosoftOfficeIntrupt : System.Web.UI.Page
{
    System.Data.DataTable dtCustmer = new System.Data.DataTable();
     object[] query;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {


        }

    }

    private void GetRecoredForExcelfile()
    {
        using (OpenXmlDataDataContext db = new OpenXmlDataDataContext())
        {
            var info = from p in db.userinfos
                       select p;

            if (info != null)
            {
                query = info.ToArray();
                dtCustmer = ConvertToDatatable(query);
                //Session["dtlist"] =Dt;
            }

        }
    }


    /// <summary>
    /// Convert Object Array to DataTable
    /// </summary>
    /// <param name="array"></param>
    /// <returns></returns>
    public static System.Data.DataTable ConvertToDatatable(Object[] array)
    {

        PropertyInfo[] properties = array.GetType().GetElementType().GetProperties();
        System.Data.DataTable dt = CreateDataTable(properties);
        if (array.Length != 0)
        {
            foreach (object o in array)
                FillData(properties, dt, o);
        }
        return dt;
    }

    #region Private Methods

    /// <summary>
    /// Creates total column of datatable.
    /// </summary>
    /// <param name="properties"></param>
    /// <returns></returns>
    private static System.Data.DataTable CreateDataTable(PropertyInfo[] properties)
    {
        System.Data.DataTable dt = new System.Data.DataTable();
        DataColumn dc = null;
        foreach (PropertyInfo pi in properties)
        {
            dc = new DataColumn();
            dc.ColumnName = pi.Name;
            //dc.DataType = pi.PropertyType;
            dt.Columns.Add(dc);
        }
        return dt;
    }

    /// <summary>
    /// Fills data in Datatable
    /// </summary>
    /// <param name="properties"></param>
    /// <param name="dt"></param>        
    private static void FillData(PropertyInfo[] properties, System.Data.DataTable dt, Object o)
    {
        DataRow dr = dt.NewRow();
        foreach (PropertyInfo pi in properties)
        {
            dr[pi.Name] = pi.GetValue(o, null);
        }
        dt.Rows.Add(dr);
    }

    #endregion
    protected void btn_Excel_Click(object sender, EventArgs e)
    
    {
        GetRecoredForExcelfile();
        string newFilePath = Server.MapPath("ExcelFile/OfficeErrorList.xlsx");        
            ApplicationClass objExcel = null;
            Workbooks objBooks = null;
            _Workbook objBook = null;
            Sheets objSheets = null;
            _Worksheet objSheet = null;
            Range objRange = null;
            int row = 1, col = 1;
            try
                {
                //   System.Data.DataTable dtCustmer = GetAllCustomers();
                   //System.Data.DataTable dtCustmer = Dt.Clone();
                   objExcel = new ApplicationClass();
                   objBooks = objExcel.Workbooks;
                   objBook = objBooks.Add(XlWBATemplate.xlWBATWorksheet);
                    //Print column heading in the excel sheet
                    int j = col;
                    foreach (DataColumn column in dtCustmer.Columns)
                        {
                            objSheets = objBook.Worksheets;
                            objSheet = (_Worksheet)objSheets.get_Item(1);
                            objRange = (Range)objSheet.Cells[row, j];
                            objRange.Value2 = column.ColumnName;
                           // objRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                            //objRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon);
                            j++;
                        }
                        row++;

                        int count = dtCustmer.Columns.Count;
                        foreach (DataRow dataRow in dtCustmer.Rows)
                        {
                            int k = col;
                            for (int i = 0; i < count; i++)
                            {
                                objRange = (Range)objSheet.Cells[row, k];
                                objRange.Value2 = dataRow[i].ToString();
                                k++;
                            }
                        row++;
                        }

                        //Save Excel document
                        objSheet.Name = "Sample Sheet";
                        object objOpt = Missing.Value;
                        objBook.SaveAs(newFilePath, objOpt, objOpt, objOpt, objOpt, objOpt, XlSaveAsAccessMode.xlNoChange, objOpt, objOpt, objOpt, objOpt, objOpt);
                        objBook.Close(false, objOpt, objOpt);
                    
                }
            catch
                {
                }
            finally
                {
                    objExcel = null;
                    objBooks = null;
                    objBook = null;
                    objSheets = null;
                    objSheet = null;
                    objRange = null;
                    ReleaseComObject(objExcel);
                    ReleaseComObject(objBooks);
                    ReleaseComObject(objBook);
                    ReleaseComObject(objSheets);
                    ReleaseComObject(objSheet);
                    ReleaseComObject(objRange);
                }


        }
     //Release COM objects from memory
    public void ReleaseComObject(object reference)
    {
        try
        {
            while (System.Runtime.InteropServices.Marshal.ReleaseComObject(reference) <= 0)
            {
            }
        }
        catch
        {
        }
    }    
}