using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Formatting.FormatStrings;
using Telerik.Windows.Documents.Spreadsheet.Model;

namespace excel2
{
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        public RadForm1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            XlsxFormatProvider formatProvider = new XlsxFormatProvider();
            Workbook workbook = formatProvider.Import(File.ReadAllBytes(@"C:\datos\rascon\datos.xlsx"));

            var worksheet = workbook.Sheets[0] as Worksheet;
            var table = new DataTable();


            //for (int i = 0; i < worksheet.UsedCellRange.ColumnCount; i++)
            //{
            //    CellSelection selection = worksheet.Cells[0, i];
            //    var columnName = selection.GetValue().Value.RawValue.ToString();

            //    table.Columns.Add(columnName);
            //}
            table.Columns.Add("I");
            table.Columns.Add("CUIT");
            table.Columns.Add("Agente");
            table.Columns.Add("Mes");
            table.Columns.Add("Año");
            table.Columns.Add("DDJJ");
            table.Columns.Add("Rect");
            table.Columns.Add("Comprobante");
            table.Columns.Add("Tipo Comprobante");
            table.Columns.Add("Régimen");
            table.Columns.Add("Monto Sujeto");
            table.Columns.Add("Alícuota");
            table.Columns.Add("Monto Retenido");
            table.Columns.Add("Fecha Retención");
            table.Columns.Add("Tipo Operación");
            table.Columns.Add("Fecha Constancia");
            table.Columns.Add("DDNro ConstanciaJJ");
            table.Columns.Add("Nro Constancia Original");
            table.Columns.Add("F");


            for (int i = 3; i < worksheet.UsedCellRange.RowCount; i++)
            {
                var values = new object[worksheet.UsedCellRange.ColumnCount];

                for (int j = 0; j < worksheet.UsedCellRange.ColumnCount; j++)
                {
                    CellSelection selection = worksheet.Cells[i, j];

                    ICellValue value = selection.GetValue().Value;
                    CellValueFormat format = selection.GetFormat().Value;
                    CellValueFormatResult formatResult = format.GetFormatResult(value);
                    string result = formatResult.InfosText;

                    values[j] = result;
                }
                table.Rows.Add(values);

            }
            //recorro otro excel
            formatProvider = new XlsxFormatProvider();
            workbook = formatProvider.Import(File.ReadAllBytes(@"C:\datos\rascon\datos2.xlsx"));
            worksheet = workbook.Sheets[0] as Worksheet;

            for (int i = 3; i < worksheet.UsedCellRange.RowCount; i++)
            {
                var values = new object[worksheet.UsedCellRange.ColumnCount];

                for (int j = 0; j < worksheet.UsedCellRange.ColumnCount; j++)
                {
                    CellSelection selection = worksheet.Cells[i, j];

                    ICellValue value = selection.GetValue().Value;
                    CellValueFormat format = selection.GetFormat().Value;
                    CellValueFormatResult formatResult = format.GetFormatResult(value);
                    string result = formatResult.InfosText;

                    values[j] = result;
                }
                table.Rows.Add(values);

            }


            radGridView1.DataSource = table;
            // Step 1: Convert a DataTable to Workbook
            DataTableFormatProvider provider = new DataTableFormatProvider();

            Workbook workbook2 = new Workbook();
            Worksheet worksheet2 = workbook2.Worksheets.Add();

            provider.Import(table, worksheet2);

            // Step 2: Save Workbook as Excel file
            IWorkbookFormatProvider formatProvider2 = new XlsxFormatProvider();

            using (Stream output = new FileStream("union.xlsx", FileMode.Create))
            {
                formatProvider2.Export(workbook2, output);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           

        }
    }
}
