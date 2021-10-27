using DevExpress.DataAccess.Excel;
using DevExpress.SpreadsheetSource;
using DevExpress.Utils.DragDrop;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
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
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace LianGangPinGuan_DoThi
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
            HandleBehaviorDragDropEvents();

        }
        ExcelDataSource ds1;
        DataTable ds = new DataTable();
        DataTable tbl_LocXong = new DataTable();
        DataTable tbl_LineChat;
        DataTable tblChart = new DataTable();

        bool isBam1 = false;
        bool isBam2 = false;

        // Đọc file ini
        iniFile iniFile = new iniFile(Application.StartupPath + "\\link.ini");

        //giải mã
        public static string Decrypt(string encodedText, string key)
        {
            TripleDESCryptoServiceProvider desCryptoProvider = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider hashMD5Provider = new MD5CryptoServiceProvider();

            byte[] byteHash;
            byte[] byteBuff;

            byteHash = hashMD5Provider.ComputeHash(Encoding.UTF8.GetBytes(key));
            desCryptoProvider.Key = byteHash;
            desCryptoProvider.Mode = CipherMode.ECB; //CBC, CFB
            byteBuff = Convert.FromBase64String(encodedText);

            string plaintext = Encoding.UTF8.GetString(desCryptoProvider.CreateDecryptor().TransformFinalBlock(byteBuff, 0, byteBuff.Length));
            return plaintext;
        }

        //Check ECB
        private void Form1_Load(object sender, EventArgs e)
        {
            Int64 ECBJob = Convert.ToInt64(Decrypt(iniFile.Read("ECB", "Link"), "LAMSON"));
            if (ECBJob.ToECB())
            {
                MessageBox.Show("ExcelDataReader.Dataset.dll.ECB cann't read data.AppConfig.For instruction on using this program, refer to its Help file.");
                return;
            }
        }

        string filePath = "";

        //Convert excel to ExcelDAtaSource :Lấy kiểu này để lấy được type của từng cột trong file excel
        private ExcelDataSource GetExcelDataSource()
        {
            ExcelDataSource ds = new ExcelDataSource();
            ds.FileName = filePath;
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
            using (ISpreadsheetSource spreadsheetSource = SpreadsheetSourceFactory.CreateSource(filePath))
            {
                IWorksheetCollection worksheetCollection = spreadsheetSource.Worksheets;
                worksheetName = worksheetCollection[p].Name;
            }
            return worksheetName;
        }

        //convert ExcelDataSource to DataTable (DevExpress)
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
                filePath = odf.FileName;
                txtPath.Text = odf.SafeFileName;
            }
            else
            {
                XtraMessageBox.Show("Chưa chọn file!");
                txtPath.Text = "";
                return;
            }

            SplashScreenManager.ShowDefaultWaitForm();

            ds1 = GetExcelDataSource(); // excel file to excel datasource
            ds = ToDataTable(ds1);  // Excel datasourcer to DataTable
            pivotGridControl1.DataSource = ds1;  //add ExcelDatasourcer to PivotGridcontrol

            List<string> LstColumn = new List<string>();

            for (int i = 0; i < ds.Columns.Count; i++)
            {
                LstColumn.Add(ds.Columns[i].ColumnName);
            }

            Check_Column.Properties.Items.AddRange(LstColumn.ToArray()); // add list to DEV CheckedCombobox

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

            tbl_LocXong.Clear();// xóa dữ liệu trong table

            if (txtPath.Text == "")
            {
                XtraMessageBox.Show("Xin vui lòng nhập đường dẫn!");
                return;
            }
            isBam1 = true;
            pivotGridControl1.Fields.Clear();

            //add filter to PivotGridcontrol
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

            //add row filter
            PivotGridField hang_Row = new PivotGridField()
            {
                Area = PivotArea.RowArea,
                AreaIndex = 0,
                Caption = cbbTrucCot.Text,
                FieldName = cbbTrucCot.Text
            };
            pivotGridControl1.Fields.Add(hang_Row);

            //add column filter
            PivotGridField cot = new PivotGridField()
            {
                Area = PivotArea.ColumnArea,
                AreaIndex = 0,
                Caption = cbbTrucDong.Text,
                //GroupInterval = PivotGroupInterval.Alphabetical,
                FieldName = cbbTrucDong.Text
            };
            pivotGridControl1.Fields.Add(cot);


            //add data filter
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

            // chuyển Pivot sau khi lọc xong thành datatable ( table theo kiểu định dạng của Pivot)
            PivotSummaryDataSource ds2 = pivotGridControl1.CreateSummaryDataSource();
            tbl_LocXong = ConvertPivotSummaryDataSourceToDataTable(ds2);

            for (int i = 0; i < tbl_LocXong.Rows.Count; i++)
            {
                if (string.IsNullOrEmpty(tbl_LocXong.Rows[i][0].ToString()))
                {
                    tbl_LocXong.Rows[i][0] = "";
                }

                if (tbl_LocXong.Rows[i][1].ToString().Replace(" ", "") == "")
                {
                    tbl_LocXong.Rows[i][1] = "NA";
                }

            }
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
            //xóa dữ liệu của lần vẽ trước đó
            chartControl1.SeriesDataMember = null;
            chartControl1.SeriesTemplate.ArgumentDataMember = null;
            chartControl1.DataSource = null;
            chartControl1.Series.Clear();
            tblChart.Clear();

            if (!isBam2)
            {
                XtraMessageBox.Show("Xin vui lòng bấm bước 2!");
                return;
            }

            if (cbbLoaiDoThi.Text == "")
            {
                XtraMessageBox.Show("Xin vui lòng chọn loại đồ thị!");
                return;
            }

            try
            {
                tbl_LocXong.Columns.Add("BieuDoChong", typeof(double));
                tbl_LocXong.Columns.Add("BieuDoPhanTram", typeof(double));
            }
            catch (Exception)
            {

            }

            //tính value dạng cột chồng
            for (int i = 0; i < tbl_LocXong.Rows.Count; i++)
            {
                tbl_LocXong.Rows[i][3] = 0;
                string Column1_Name = tbl_LocXong.Columns[1].ColumnName;
                string Column2_Name = tbl_LocXong.Columns[2].ColumnName;
                string Value_Seach = tbl_LocXong.Rows[i][1].ToString();

                int SUM = tbl_LocXong.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<int>(Column2_Name));

                tbl_LocXong.Rows[i][3] = Math.Round((Convert.ToDouble(tbl_LocXong.Rows[i][2]) / SUM) * 100, 2);
            }

            tbl_LocXong.AcceptChanges();

            var Query = from tbl1 in tbl_LocXong.AsEnumerable()
                        join tbl2 in tbl_LineChat.AsEnumerable() on tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName).ToString() equals tbl2.Field<string>("Line")
                    // orderby (tbl1.Field<object>(tbl_LocXong.Columns[1].ColumnName))
                    //orderby (tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName))
                        select tbl1;

            //var Query = from tbl1 in tbl_LocXong.AsEnumerable()
            //            join tbl2 in tbl_LineChat.AsEnumerable() on tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName).ToString() equals tbl2.Field<string>("Line")
            //          //  orderby (tbl1.Field<object>(tbl_LocXong.Columns[1].ColumnName))
            //           // orderby (tbl1.Field<object>(tbl_LocXong.Columns[0].ColumnName))
            //           select tbl1;
           
            try
            {
                tblChart = Query.CopyToDataTable();
            }
            catch (Exception)
            {

                XtraMessageBox.Show("Vui lòng chọn bước 2, sau đó vẽ đồ thị");
                return;
            }


            // Tính value dạng biểu đồ phần trăm
            for (int i = 0; i < tblChart.Rows.Count; i++)
            {
                tblChart.Rows[i][4] = 0;
                string Column1_Name = tblChart.Columns[1].ColumnName;
                string Column2_Name = tblChart.Columns[2].ColumnName;
                string Value_Seach = tblChart.Rows[i][1].ToString();

                int SUM = tblChart.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<int>(Column2_Name));

                tblChart.Rows[i][4] = Math.Round((Convert.ToDouble(tblChart.Rows[i][2]) / SUM) * 100, 2);
            }

            var ww1 = tblChart.AsEnumerable().OrderBy(r => r.Field<object>(tblChart.Columns[0].ColumnName)).ThenBy(r => r.Field<object>(tblChart.Columns[1].ColumnName));

            DataTable tblChart1 = ww1.CopyToDataTable();
            //DataView dv = tblChart.DefaultView;
            // dv.Sort = $"{tblChart.Columns[1].ColumnName} ASC";
            //tblChart = dv.ToTable();

            chartControl1.DataSource = tblChart1;



            // Specify data members to bind the chart's series template. (3 yếu tố tạo thành 1 đồ thị: tên các đường +Trục X +Value)
            chartControl1.SeriesDataMember = tbl_LocXong.Columns[0].ColumnName;  //SeriesName
            chartControl1.SeriesTemplate.ArgumentDataMember = tbl_LocXong.Columns[1].ColumnName; // Trục X

            // Chọn kiểu đồ thị
            if (cbbLoaiDoThi.Text == "百分比堆疊直條圖")
            {
                chartControl1.SeriesTemplate.ValueDataMembers.AddRange(new string[] { tbl_LocXong.Columns[4].ColumnName }); //Values
                chartControl1.SeriesTemplate.View = new FullStackedBarSeriesView();
                ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextPattern = "{V:p0}";         
            }
            else
            {
                chartControl1.SeriesTemplate.ValueDataMembers.AddRange(new string[] { tbl_LocXong.Columns[3].ColumnName });
                chartControl1.SeriesTemplate.View = new StackedBarSeriesView();
                ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextPattern = "{V:F2} %";
            }

            // Làm đẹp cho đồ thị
            ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextColor = Color.Red;
            ((XYDiagram)chartControl1.Diagram).AxisX.Label.TextPattern = "{V:F2}";
            ((XYDiagram)chartControl1.Diagram).AxisX.Label.TextColor = Color.Black;
            ((XYDiagram)chartControl1.Diagram).EnableAxisXZooming = true;
            ((XYDiagram)chartControl1.Diagram).EnableAxisXScrolling = true;
            ((XYDiagram)chartControl1.Diagram).ZoomingOptions.AxisXMaxZoomPercent = 200;
            chartControl1.Legend.Font = Font = new Font("DFKai-SB", 12, FontStyle.Regular);
            for (int i = 0; i < tbl_LineChat.Rows.Count; i++)
            {

                //chartControl1.Series[i].ArgumentScaleType = ScaleType.Auto;

                //((StackedBarSeriesView)chartControl1.Series[i].View).MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
                //((StackedBarSeriesView)chartControl1.Series[i].View).LineMarkerOptions.Kind = MarkerKind.Circle;
                if (check_HienThiSoLieu.Checked == true)
                    chartControl1.Series[i].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;


                chartControl1.Series[i].Label.TextPattern = "{V:F2}";
                if (tbl_LineChat.Rows[i][1].ToString() != "")
                {
                    chartControl1.Series[i].View.Color = (Color)tbl_LineChat.Rows[i][1];
                }
            }
            // isBam2 = false;
        }



        private void Behavior_DragDrop(object sender, DevExpress.Utils.DragDrop.DragDropEventArgs e)
        {
            GridView targetGrid = e.Target as GridView;
            GridView sourceGrid = e.Source as GridView;
            if (e.Action == DragDropActions.None || targetGrid != sourceGrid)
                return;
            DataTable sourceTable = sourceGrid.GridControl.DataSource as DataTable;

            Point hitPoint = targetGrid.GridControl.PointToClient(Cursor.Position);
            GridHitInfo hitInfo = targetGrid.CalcHitInfo(hitPoint);

            int[] sourceHandles = e.GetData<int[]>();

            int targetRowHandle = hitInfo.RowHandle;
            int targetRowIndex = targetGrid.GetDataSourceRowIndex(targetRowHandle);

            List<DataRow> draggedRows = new List<DataRow>();
            foreach (int sourceHandle in sourceHandles)
            {
                int oldRowIndex = sourceGrid.GetDataSourceRowIndex(sourceHandle);
                DataRow oldRow = sourceTable.Rows[oldRowIndex];
                draggedRows.Add(oldRow);
            }

            int newRowIndex;

            switch (e.InsertType)
            {
                case InsertType.Before:
                    newRowIndex = targetRowIndex > sourceHandles[sourceHandles.Length - 1] ? targetRowIndex - 1 : targetRowIndex;
                    for (int i = draggedRows.Count - 1; i >= 0; i--)
                    {
                        DataRow oldRow = draggedRows[i];
                        DataRow newRow = sourceTable.NewRow();
                        newRow.ItemArray = oldRow.ItemArray;
                        sourceTable.Rows.Remove(oldRow);
                        sourceTable.Rows.InsertAt(newRow, newRowIndex);
                    }
                    break;
                case InsertType.After:
                    newRowIndex = targetRowIndex < sourceHandles[0] ? targetRowIndex + 1 : targetRowIndex;
                    for (int i = 0; i < draggedRows.Count; i++)
                    {
                        DataRow oldRow = draggedRows[i];
                        DataRow newRow = sourceTable.NewRow();
                        newRow.ItemArray = oldRow.ItemArray;
                        sourceTable.Rows.Remove(oldRow);
                        sourceTable.Rows.InsertAt(newRow, newRowIndex);
                    }
                    break;
                default:
                    newRowIndex = -1;
                    break;
            }
            int insertedIndex = targetGrid.GetRowHandle(newRowIndex);
            targetGrid.FocusedRowHandle = insertedIndex;
            targetGrid.SelectRow(targetGrid.FocusedRowHandle);
        }

        private void Behavior_DragOver(object sender, DragOverEventArgs e)
        {
            DragOverGridEventArgs args = DragOverGridEventArgs.GetDragOverGridEventArgs(e);
            e.InsertType = args.InsertType;
            e.InsertIndicatorLocation = args.InsertIndicatorLocation;
            e.Action = args.Action;
            Cursor.Current = args.Cursor;
            args.Handled = true;
        }

        public void HandleBehaviorDragDropEvents()
        {
            DragDropBehavior gridControlBehavior = behaviorManager1.GetBehavior<DragDropBehavior>(this.gridView1);
            gridControlBehavior.DragDrop += Behavior_DragDrop;
            gridControlBehavior.DragOver += Behavior_DragOver;
        }

        private void btnXuatExcel_Click(object sender, EventArgs e)
        {
    //        var query = tblChart.AsEnumerable()
    //.GroupBy(row => row.Field<object>(tblChart.Columns[0].ColumnName))
    //.Select(g => new {
    //    ColumnName = g.Key,
    //    a = g.Where(row => row.Field<string>("Property") == "a").Select(c => c.Field<string>("Value")),
    //    b = g.Where(row => row.Field<string>("Property") == "b").Select(c => c.Field<string>("Value")),
    //    c = g.Where(row => row.Field<string>("Property") == "c").Select(c => c.Field<string>("Value")),
    //    d = g.Where(row => row.Field<string>("Property") == "d").Select(c => c.Field<string>("Value"))
    //});
  }
    }
}

