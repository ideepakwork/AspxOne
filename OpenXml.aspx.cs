using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Reflection;
using System.IO;
using OfficeOpenXml;

public partial class OpenXml : System.Web.UI.Page
{
    DataTable Dt = new DataTable();
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
                Dt = ConvertToDatatable(query);            
            }

        }
    }


    /// <summary>
    /// Convert Object Array to DataTable
    /// </summary>
    /// <param name="array"></param>
    /// <returns></returns>
    public static DataTable ConvertToDatatable(Object[] array)
    {

        PropertyInfo[] properties = array.GetType().GetElementType().GetProperties();
        DataTable dt = CreateDataTable(properties);
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
    private static DataTable CreateDataTable(PropertyInfo[] properties)
    {
        DataTable dt = new DataTable();
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
    private static void FillData(PropertyInfo[] properties, DataTable dt, Object o)
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
        string newFilePath = Server.MapPath("ExcelFile/ErrorList.xlsx");
        string templateFilePath = Server.MapPath("ExcelFile/ErrorListtemplate.xlsx");
        FileInfo newFile = new FileInfo(newFilePath);
        FileInfo template = new FileInfo(templateFilePath);
        using (ExcelPackage xlPackage = new ExcelPackage(newFile, template))
        {
            
            foreach (ExcelWorksheet aworksheet in xlPackage.Workbook.Worksheets)
            {
                aworksheet.Cell(1, 1).Value = aworksheet.Cell(1, 1).Value;
            }

            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Sheet1"];
            int startrow = 5;
            int row = 0;
            int col = 0;

            for (int j = 0; j < Dt.Columns.Count; j++)
            {
               // col = col + j;
                col++;
                for (int i = 0; i < Dt.Rows.Count; i++)
                {
                    row = startrow + i;                   
                    ExcelCell cell = worksheet.Cell(row, col);
                    cell.Value = Dt.Rows[i][j].ToString();
                    xlPackage.Save();
                }
            } 
        }
    }
}