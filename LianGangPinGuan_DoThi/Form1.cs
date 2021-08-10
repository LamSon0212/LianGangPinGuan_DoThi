using DevExpress.DataAccess.Excel;
using DevExpress.SpreadsheetSource;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraSplashScreen;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LianGangPinGuan_DoThi
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();

        }
        ExcelDataSource ds1;
        DataTable ds = new DataTable();
        DataTable tbl_LocXong = new DataTable();
        DataTable tbl_LineChat;

        bool isBam1 = false;
        bool isBam2 = false;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        string fileName = "";

        private ExcelDataSource GetExcelDataSource()
        {
            ExcelDataSource ds = new ExcelDataSource();
            ds.FileName = fileName;
            DevExpress.DataAccess.Excel.ExcelSourceOptions excelSourceOptions1 = new DevExpress.DataAccess.Excel.ExcelSourceOptions();
            DevExpress.DataAccess.Excel.ExcelWorksheetSettings excelWorksheetSettings1 = new DevExpress.DataAccess.Excel.ExcelWorksheetSettings();
            excelWorksheetSettings1.WorksheetName = GetWorkSheetNameByIndex(0);
            excelSourceOptions1.ImportSettings = excelWorksheetSettings1;
            ds.SourceOptions = excelSourceOptions1;
            ds.Fill();
            return ds;
        }

        private string GetWorkSheetNameByIndex(int p)
        {
            string worksheetName = "";
            using (ISpreadsheetSource spreadsheetSource = SpreadsheetSourceFactory.CreateSource(fileName))
            {
                IWorksheetCollection worksheetCollection = spreadsheetSource.Worksheets;
                worksheetName = worksheetCollection[p].Name;
            }
            return worksheetName;
        }

        public DataTable ToDataTable(ExcelDataSource excelDataSource)
        {
            IList list = ((IListSource)excelDataSource).GetList();
            DevExpress.DataAccess.Native.Excel.DataView dataView = (DevExpress.DataAccess.Native.Excel.DataView)list;
            List<PropertyDescriptor> props = dataView.Columns.ToList<PropertyDescriptor>();
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);
            }
            object[] values = new object[props.Count];
            foreach (DevExpress.DataAccess.Native.Excel.ViewRow item in list)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }

        private DataTable ConvertPivotSummaryDataSourceToDataTable(PivotSummaryDataSource source)
        {
            DataTable dt = new DataTable();
            foreach (PropertyDescriptor _propertyDescriptor in source.GetItemProperties(null))
                dt.Columns.Add(_propertyDescriptor.Name, _propertyDescriptor.PropertyType);
            for (int r = 0; r < source.RowCount; r++)
            {
                object[] rowValues = new object[dt.Columns.Count];
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    rowValues[c] = source.GetValue(r, dt.Columns[c].ColumnName);
                }
                dt.Rows.Add(rowValues);
            }
            return dt;
        }


        private void txtPath_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog odf = new OpenFileDialog();
            odf.Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls";
            if (odf.ShowDialog() == DialogResult.OK)
            {
                fileName = odf.FileName;
                txtPath.Text = odf.SafeFileName;
            }
            else
            {
                XtraMessageBox.Show("Chưa chọn file!");
                txtPath.Text = "";
                return;
            }

            SplashScreenManager.ShowDefaultWaitForm();

            ds1 = GetExcelDataSource();
            ds = ToDataTable(ds1);
            pivotGridControl1.DataSource = ds1;

            List<string> LstColumn = new List<string>();

            for (int i = 0; i < ds.Columns.Count; i++)
            {
                LstColumn.Add(ds.Columns[i].ColumnName);
            }

            Check_Column.Properties.Items.AddRange(LstColumn.ToArray());

            SplashScreenManager.CloseDefaultSplashScreen();
        }


        // add Item to CbbTrucX, CbbTrucY
        private void Check_Column_EditValueChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < Check_Column.Properties.Items.Count; i++)
            {
                if (Check_Column.Properties.Items[i].CheckState == CheckState.Unchecked)
                {
                    cbbTrucCot.Properties.Items.Add(Check_Column.Properties.Items[i].ToString());
                    cbbValue.Properties.Items.Add(Check_Column.Properties.Items[i].ToString());
                }

            }
        }

        private void cbbTrucX_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbbTrucDong.Properties.Items.Clear();
            cbbTrucDong.Properties.Items.AddRange(cbbTrucCot.Properties.Items);
            cbbTrucDong.Properties.Items.Remove(cbbTrucCot.Text);
        }


        private void txtPath_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void btnXuLiDieuKien_Click(object sender, EventArgs e)
        {
            if (txtPath.Text=="")
            {
                XtraMessageBox.Show("Xin vui lòng nhập đường dẫn!");
                return;
            }
            isBam1 = true;
            pivotGridControl1.Fields.Clear();
            for (int i = 0; i < Check_Column.Properties.Items.Count; i++)
            {
                if (Check_Column.Properties.Items[i].CheckState == CheckState.Checked)
                {
                    PivotGridField filter = new PivotGridField()
                    {

                        AreaIndex = i,
                        Caption = Check_Column.Properties.Items[i].ToString(),
                        FieldName = Check_Column.Properties.Items[i].ToString()
                    };

                    pivotGridControl1.Fields.Add(filter);
                }
            }

            PivotGridField hang_Row = new PivotGridField()
            {
                Area = PivotArea.RowArea,
                AreaIndex = 0,
                Caption = cbbTrucCot.Text,
                FieldName = cbbTrucCot.Text
            };
            pivotGridControl1.Fields.Add(hang_Row);

            PivotGridField cot = new PivotGridField()
            {
                Area = PivotArea.ColumnArea,
                AreaIndex = 0,
                Caption = cbbTrucDong.Text,
                //GroupInterval = PivotGroupInterval.Alphabetical,
                FieldName = cbbTrucDong.Text
            };
            pivotGridControl1.Fields.Add(cot);

            PivotGridField data1 = new PivotGridField()
            {
                Area = PivotArea.DataArea,
                SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Count,
                AreaIndex = 0,
                Caption = cbbValue.Text,
                FieldName = cbbValue.Text
            };
            pivotGridControl1.Fields.Add(data1);

        }

        private void btnTaoDuongChart_Click(object sender, EventArgs e)
        {
            if (!isBam1)
            {
                XtraMessageBox.Show("Xin vui lòng bấm bước 1!");
                return;
            }
            PivotSummaryDataSource ds2 = pivotGridControl1.CreateSummaryDataSource();
            tbl_LocXong = ConvertPivotSummaryDataSourceToDataTable(ds2);

            //xử lí bảng để lấy các đường
            var list_Cot = tbl_LocXong.AsEnumerable().Select(r => r.Field<object>(cbbTrucCot.Text)).Distinct().ToList();

            tbl_LineChat = new DataTable();
            tbl_LineChat.Columns.Add(new DataColumn("Line", typeof(string)));
            tbl_LineChat.Columns.Add(new DataColumn("ColorLine", typeof(Color)));
            for (int i = 0; i < list_Cot.Count; i++)
            {
                tbl_LineChat.Rows.Add(list_Cot[i]);
            }
            gridControl1.DataSource = tbl_LineChat;

            isBam2 = true;
        }

        private void btnXoa_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            int indexR = gridView1.FocusedRowHandle;
            tbl_LineChat.Rows.RemoveAt(indexR);
            isBam2 = true;

        }

        private void btnVeDoThi_Click(object sender, EventArgs e)
        {
            chartControl1.DataSource = null;
            chartControl1.Series.Clear();
            if (!isBam2)
            {
                XtraMessageBox.Show("Xin vui lòng bấm bước 2!");
                return;
            }

            try
            {
                tbl_LocXong.Columns.Add("SUM", typeof(double));
            }
            catch (Exception)
            {
            }


            for (int i = 0; i < tbl_LocXong.Rows.Count; i++)
            {
                tbl_LocXong.Rows[i][3] = 0;
                string c = tbl_LocXong.Columns[1].ColumnName;
                string b = tbl_LocXong.Columns[2].ColumnName;
                string a = tbl_LocXong.Rows[i][1].ToString();

                int sum = tbl_LocXong.AsEnumerable().Where(row => row.Field<object>(c).ToString() == a).Sum(row => row.Field<int>(b));
                tbl_LocXong.Rows[i][3] = Math.Round((Convert.ToDouble(tbl_LocXong.Rows[i][2]) / sum)*100,2);
            }

            tbl_LocXong.AcceptChanges();

            var Query = from tbl1 in tbl_LocXong.AsEnumerable()
                        join tbl2 in tbl_LineChat.AsEnumerable() on tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName).ToString() equals tbl2.Field<string>("Line")
                        orderby (tbl1.Field<object>(tbl_LocXong.Columns[1].ColumnName))
                        orderby (tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName))
                        select tbl1;

            DataTable tblChart = Query.CopyToDataTable();

            var ww1 = tblChart.AsEnumerable().OrderBy(r => r.Field<object>(tblChart.Columns[0].ColumnName)).ThenBy(r => r.Field<object>(tblChart.Columns[1].ColumnName));

            DataTable tblChart1 = ww1.CopyToDataTable();
            //DataView dv = tblChart.DefaultView;
            // dv.Sort = $"{tblChart.Columns[1].ColumnName} ASC";
            //tblChart = dv.ToTable();

            chartControl1.DataSource = tblChart1;

            

            // Specify data members to bind the chart's series template.
            chartControl1.SeriesDataMember = tbl_LocXong.Columns[0].ColumnName;
            chartControl1.SeriesTemplate.ArgumentDataMember = tbl_LocXong.Columns[1].ColumnName;
            chartControl1.SeriesTemplate.ValueDataMembers.AddRange(new string[] { tbl_LocXong.Columns[3].ColumnName });

            // Specify the template's series view.
            chartControl1.SeriesTemplate.View = new FullStackedBarSeriesView();

            // Làm đẹp cho đồ thị
            ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextPattern = "{V:p0}";
            ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextColor = Color.Red;
            ((XYDiagram)chartControl1.Diagram).AxisX.Label.TextPattern = "{V:F2}";
            ((XYDiagram)chartControl1.Diagram).AxisX.Label.TextColor = Color.Black;
            ((XYDiagram)chartControl1.Diagram).EnableAxisXZooming = true;
            ((XYDiagram)chartControl1.Diagram).EnableAxisXScrolling = true;
            ((XYDiagram)chartControl1.Diagram).ZoomingOptions.AxisXMaxZoomPercent = 200;
            chartControl1.Legend.Font = Font = new Font("DFKai-SB", 12, FontStyle.Regular);
            for (int i = 0; i < tbl_LineChat.Rows.Count; i++)
            {
                //((StackedBarSeriesView)chartControl1.Series[i].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
                //((StackedBarSeriesView)chartControl1.Series[i].View).LineMarkerOptions.Kind = MarkerKind.Circle;
                chartControl1.Series[i].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                chartControl1.Series[i].Label.TextPattern = "{V:F2}";
                if (tbl_LineChat.Rows[i][1].ToString() != "")
                {
                    chartControl1.Series[i].View.Color = (Color)tbl_LineChat.Rows[i][1];
                }


            }

            isBam2 = false;

            // dtResult = customerNames.CopyToDataTable();

            //dtResult.Columns.Add("SeriesName", typeof(string));
            //dtResult.Columns.Add("Cot", typeof(string));
            //dtResult.Columns.Add("Value", typeof(string));

            //var result = from dataRows1 in tbl_LocXong.AsEnumerable()
            //             join dataRows2 in tbl_LineChat.AsEnumerable()
            //             on dataRows1.Field<object>(tbl_LocXong.Columns[0].ColumnName) equals dataRows2.Field<string>("Line")

            //             select dtResult.LoadDataRow(new object[]
            //             {
            //    dataRows1.Field<string>(tbl_LocXong.Columns[0].ColumnName),
            //    dataRows1.Field<string>(tbl_LocXong.Columns[1].ColumnName),
            //    dataRows1.Field<string>(tbl_LocXong.Columns[2].ColumnName),
            //              }, false);
            //DataTable b = result.CopyToDataTable();

            //var query = from order in tbl_LocXong.AsEnumerable()
            //            join detail in tbl_LineChat.AsEnumerable()
            //            on order.Field<int>("SalesOrderID") equals
            //                detail.Field<int>("SalesOrderID")
            //            where order.Field<bool>("OnlineOrderFlag") == true
            //            && order.Field<DateTime>("OrderDate").Month == 8
            //            select new
            //            {
            //                SalesOrderID =
            //                    order.Field<int>("SalesOrderID"),
            //                SalesOrderDetailID =
            //                    detail.Field<int>("SalesOrderDetailID"),
            //                OrderDate =
            //                    order.Field<DateTime>("OrderDate"),
            //                ProductID =
            //                    detail.Field<int>("ProductID")
            //            };

            //DataTable orderTable = query.CopyToDataTable();

        }

        private void chartControl1_CustomDrawCrosshair(object sender, CustomDrawCrosshairEventArgs e)
        {
            foreach (CrosshairElement element in e.CrosshairElements)
            {
                element.LabelElement.Font = new Font("DFKai-SB", 10, FontStyle.Italic);
                element.LabelElement.TextColor = Color.Red;
                
            }
        }
    }
}
