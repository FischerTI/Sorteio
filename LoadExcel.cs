using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sorteio
{
  class LoadExcel
  {
    public LoadExcel(string path)
    {
      if (!string.IsNullOrEmpty(path))
      {
        ReadExcelAsync(path);
      }
    }

    public List<DataTable> Tables { get; set; }

    private void ReadExcelAsync(string path)
    {
      Tables = new List<DataTable>();
      try
      {
        bool hasHeader = true;
        using (var pck = new OfficeOpenXml.ExcelPackage())
        {
          using (var stream = File.OpenRead(path))
          {
            pck.Load(stream);
          }
          foreach (var ws in pck.Workbook.Worksheets)
          {
            DataTable tbl = new DataTable(ws.Name);
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
              if (tbl.Columns.Contains(firstRowCell.Text))
                tbl.Columns.Add(hasHeader ? firstRowCell.Text + new Random(20).Next() : string.Format("Column {0}", firstRowCell.Start.Column));
              else
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
              var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
              DataRow row = tbl.Rows.Add();
              foreach (var cell in wsRow)
              {
                if (cell.Text != "#N/A" && row.Table.Columns.Count >= cell.Start.Column - 1)
                  row[cell.Start.Column - 1] = cell.Text;
              }
            }

            this.Tables.Add(tbl);
          }
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("Erro Excel =D." + ex.ToString());
      }
      /*
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
              foreach (var sheet in package.Workbook.Worksheets)
              {
                foreach (var table in sheet.Tables)
                {
                  DataTable dtToAdd = new DataTable();
                  foreach (var col in table.Columns)
                  {
                    dtToAdd.Columns.Add(col.Name, col.GetType());
                  }

                }

              }
            }
            */
    }

  }
}
