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
using ExcelDataReader;
using System.Diagnostics;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

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
        DataTable dtCloned = new DataTable();
        DataTable tbl_LineChat;
        DataTable tblChart = new DataTable();
        DataTable tblChart_CHUNG = new DataTable();
        DataTable tblChartOK = new DataTable();

        DataTable dtDuLieu = new DataTable();


        Int32 NoRow_ToTal;
        Int32 NoCol_ToTal;

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
            // DataSet aaaa= extodt(@"C:\Users\LamSon\Desktop\DATA\11月W1 (1).xlsx");

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


        DataSet extodt(string Path)
        {
            DataSet a;


            using (var stream = File.Open(Path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);


                a = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                reader.Close();
            }

            return a;
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
            // odf.Filter = "All files (*.*)|*.*";
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

            //// load the XLSM file
            //var converter = new GroupDocs.Conversion.Converter(filePath);
            //// set the convert options for XLS format
            //var convertOptions = converter.GetPossibleConversions()["xls"].ConvertOptions;
            //// convert to XLS format
            //converter.Convert("Data.xls", convertOptions);
            SplashScreenManager.ShowDefaultWaitForm();


            //DataSet aaaa = extodt(filePath);

            //var querya = aaaa.Tables[0].AsEnumerable();

            //// Create the Excel Data Source.
            //ExcelDataSource ds = new ExcelDataSource();
            //ds.FileName = filePath;
            //ExcelWorksheetSettings settings = new ExcelWorksheetSettings("Sheet1");
            //ds.SourceOptions = new ExcelSourceOptions(settings);
            //ds.Fill();

            // Set the pivot's data source.
            // pivotGridControl1.DataSource = querya;

            // ds = aaaa.Tables[0];  // Excel datasourcer to DataTable

            ds1 = GetExcelDataSource(); // excel file to excel datasource
            pivotGridControl1.DataSource = ds1;  //add ExcelDatasourcer to PivotGridcontrol

            DataTable dtlOK = ToDataTable(ds1);
            List<string> LstColumn = new List<string>();

            for (int i = 0; i < dtlOK.Columns.Count; i++)
            {
                LstColumn.Add(dtlOK.Columns[i].ColumnName);
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
                FieldName = cbbTrucCot.Text,
            };
            pivotGridControl1.Fields.Add(hang_Row);

            //add column filter
            PivotGridField cot = new PivotGridField()
            {
                Area = PivotArea.ColumnArea,
                AreaIndex = 0,
                Caption = cbbTrucDong.Text,
                //GroupInterval = PivotGroupInterval.Alphabetical,
                FieldName = cbbTrucDong.Text,
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
            dtDuLieu = new DataTable();

            if (!isBam1)
            {
                XtraMessageBox.Show("Xin vui lòng bấm bước 1!");
                return;
            }

            // chuyển Pivot sau khi lọc xong thành datatable ( table theo kiểu định dạng của Pivot)
            PivotSummaryDataSource ds2 = pivotGridControl1.CreateSummaryDataSource();
            tbl_LocXong = ConvertPivotSummaryDataSourceToDataTable(ds2);

            dtCloned = tbl_LocXong.Clone();
            dtCloned.Columns[0].DataType = typeof(string);
            dtCloned.Columns[1].DataType = typeof(string);
            dtCloned.Columns[2].DataType = typeof(double);
            foreach (DataRow row in tbl_LocXong.Rows)
            {
                dtCloned.ImportRow(row);
            }


            for (int i = 0; i < dtCloned.Rows.Count; i++)
            {
                if (string.IsNullOrEmpty(dtCloned.Rows[i][0].ToString()))
                {
                    dtCloned.Rows[i][0] = "";
                }

                if (dtCloned.Rows[i][1].ToString().Replace(" ", "") == "")
                {
                    dtCloned.Rows[i][1] = "NA";
                }

            }
            //for (int i = 0; i < tbl_LocXong.Rows.Count; i++)
            //{
            //    if (string.IsNullOrEmpty(tbl_LocXong.Rows[i][0].ToString()))
            //    {
            //        tbl_LocXong.Rows[i][0] = "";
            //    }

            //    if (tbl_LocXong.Rows[i][1].ToString().Replace(" ", "") == "")
            //    {
            //        tbl_LocXong.Rows[i][1] = "NA";
            //    }

            //}
            //xử lí bảng để lấy các đường


            var list_Cot = dtCloned.AsEnumerable().Select(r => r.Field<object>(cbbTrucCot.Text)).Distinct().ToList();

            tbl_LineChat = new DataTable();
            tbl_LineChat.Columns.Add(new DataColumn("Line", typeof(string)));
            tbl_LineChat.Columns.Add(new DataColumn("ColorLine", typeof(Color)));
            for (int i = 0; i < list_Cot.Count; i++)
            {
                tbl_LineChat.Rows.Add(list_Cot[i]);
            }
            gridControl1.DataSource = tbl_LineChat;

            //xử lí bảng chung


            var list_Row = dtCloned.AsEnumerable().Select(r => r.Field<object>(dtCloned.Columns[0].ColumnName)).Distinct().ToList();
            var list_Column = dtCloned.AsEnumerable().Select(r => r.Field<object>(dtCloned.Columns[1].ColumnName)).Distinct().ToList();
            NoRow_ToTal = list_Row.Count();
            NoCol_ToTal = list_Column.Count();

            dtDuLieu.Columns.Add(" ", typeof(string));
            for (int i = 0; i < NoCol_ToTal; i++)
            {
                dtDuLieu.Columns.Add(list_Column[i].ToString(), typeof(double));
            }
            for (int i = 0; i < NoRow_ToTal; i++)
            {
                dtDuLieu.Rows.Add(list_Row[i].ToString());
            }

            for (int i = 0; i < NoRow_ToTal; i++)
            {
                for (int j = 0; j < NoCol_ToTal; j++)
                {
                    try
                    {
                        var dValue = from row in dtCloned.AsEnumerable()
                                     where row.Field<string>(dtCloned.Columns[0].ColumnName) == dtDuLieu.Rows[i][0].ToString()
                                           && row.Field<string>(dtCloned.Columns[1].ColumnName) == dtDuLieu.Columns[j + 1].ColumnName
                                     select row.Field<double>(dtCloned.Columns[2].ColumnName);
                        double Value_Ok = dValue.ToList()[0];
                        dtDuLieu.Rows[i][j + 1] = Value_Ok;
                        dtDuLieu.AcceptChanges();
                    }
                    catch (Exception)
                    {
                        dtDuLieu.Rows[i][j + 1] = 0;
                        dtDuLieu.AcceptChanges();
                    }
                }
            }

            //for (int i = 0; i < dtCloned.Rows.Count; i++)
            //{
            //    dtCloned.Rows[i][3] = 0;
            //    string Column1_Name = dtCloned.Columns[1].ColumnName;
            //    string Column2_Name = dtCloned.Columns[2].ColumnName;
            //    string Value_Seach = dtCloned.Rows[i][1].ToString();

            //    double SUM = dtCloned.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<double>(Column2_Name));
            //    if (SUM == 0)
            //    {
            //        dtCloned.Rows[i][3] = 0;
            //    }
            //    else
            //    {
            //        dtCloned.Rows[i][3] = Math.Round((Convert.ToDouble(dtCloned.Rows[i][2]) / SUM), 2);
            //    }

            //}

            //dtCloned.AcceptChanges();

            //var Query = from tbl1 in dtCloned.AsEnumerable()
            //            join tbl2 in tbl_LineChat.AsEnumerable() on tbl1.Field<object>(dtCloned.Columns[0].ColumnName).ToString() equals tbl2.Field<string>("Line")
            //            select tbl1;
            //try
            //{
            //    tblChart_CHUNG = Query.CopyToDataTable();
            //}
            //catch (Exception)
            //{

            //    XtraMessageBox.Show("Vui lòng chọn bước 2, sau đó vẽ đồ thị");
            //    return;
            //}

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
                dtCloned.Columns.Add("BieuDoChong", typeof(double));
                dtCloned.Columns.Add("BieuDoPhanTram", typeof(double));
            }
            catch (Exception)
            {

            }

            //tính value dạng cột chồng
            for (int i = 0; i < dtCloned.Rows.Count; i++)
            {
                dtCloned.Rows[i][3] = 0;
                string Column1_Name = dtCloned.Columns[1].ColumnName;
                string Column2_Name = dtCloned.Columns[2].ColumnName;
                string Value_Seach = dtCloned.Rows[i][1].ToString();

                double SUM = dtCloned.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<double>(Column2_Name));
                if (SUM == 0)
                {
                    dtCloned.Rows[i][3] = 0;
                }
                else
                {
                    dtCloned.Rows[i][3] = Math.Round((Convert.ToDouble(dtCloned.Rows[i][2]) / SUM) * 100, 4);
                }

            }

            dtCloned.AcceptChanges();

            var Query = from tbl1 in dtCloned.AsEnumerable()
                        join tbl2 in tbl_LineChat.AsEnumerable() on tbl1.Field<object>(dtCloned.Columns[0].ColumnName).ToString() equals tbl2.Field<string>("Line")
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

                double SUM = tblChart.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<double>(Column2_Name));
                if (SUM == 0)
                {
                    tblChart.Rows[i][4] = 0;
                }
                else
                {
                    tblChart.Rows[i][4] = Math.Round((Convert.ToDouble(tblChart.Rows[i][2]) / SUM) * 100, 4);
                }
            }

            var ww1 = tblChart.AsEnumerable().OrderBy(r => r.Field<object>(tblChart.Columns[0].ColumnName)).ThenBy(r => r.Field<object>(tblChart.Columns[1].ColumnName));

            tblChartOK = ww1.CopyToDataTable();
            tblChartOK.Columns[2].DataType = typeof(double);

            //DataView dv = tblChart.DefaultView;
            // dv.Sort = $"{tblChart.Columns[1].ColumnName} ASC";
            //tblChart = dv.ToTable();

            chartControl1.DataSource = tblChartOK;



            // Specify data members to bind the chart's series template. (3 yếu tố tạo thành 1 đồ thị: tên các đường +Trục X +Value)
            chartControl1.SeriesDataMember = dtCloned.Columns[0].ColumnName;  //SeriesName
            chartControl1.SeriesTemplate.ArgumentDataMember = dtCloned.Columns[1].ColumnName; // Trục X

            // Chọn kiểu đồ thị
            if (cbbLoaiDoThi.Text == "百分比堆疊直條圖")
            {
                chartControl1.SeriesTemplate.ValueDataMembers.AddRange(new string[] { dtCloned.Columns[4].ColumnName }); //Values
                chartControl1.SeriesTemplate.View = new FullStackedBarSeriesView();
                ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextPattern = "{V:p0}";
            }
            else
            {
                chartControl1.SeriesTemplate.ValueDataMembers.AddRange(new string[] { dtCloned.Columns[3].ColumnName });
                chartControl1.SeriesTemplate.View = new StackedBarSeriesView();
                ((XYDiagram)chartControl1.Diagram).AxisY.Label.TextPattern = "{V:F2} %";
            }

            // Làm đẹp cho đồ thị
            if (Convert.ToDouble(SpinMaximum.Value)!=0)
            {
                ((XYDiagram)chartControl1.Diagram).AxisY.WholeRange.MaxValue = Convert.ToDouble(SpinMaximum.Value);
            }
           
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

        private void btnXuatExcel_Click_1(object sender, EventArgs e)
        {
            DataTable dtDuLieu2 = new DataTable();
            DataTable dtDuLieu3 = new DataTable();

            var list_Row = tblChartOK.AsEnumerable().Select(r => r.Field<object>(tblChartOK.Columns[0].ColumnName)).Distinct().ToList();
            var list_Column = tblChartOK.AsEnumerable().Select(r => r.Field<object>(tblChartOK.Columns[1].ColumnName)).Distinct().ToList();
            Int32 NoRow = list_Row.Count();
            Int32 NoCol = list_Column.Count();


            //dtDuLieu.Columns.Add(" ", typeof(string));
            dtDuLieu2.Columns.Add(" ", typeof(string));
            dtDuLieu3.Columns.Add(" ", typeof(string));
            for (int i = 0; i < NoCol; i++)
            {
                // dtDuLieu.Columns.Add(list_Column[i].ToString(), typeof(double));
                dtDuLieu2.Columns.Add(list_Column[i].ToString(), typeof(double));
                dtDuLieu3.Columns.Add(list_Column[i].ToString(), typeof(double));
            }
            for (int i = 0; i < NoRow; i++)
            {
                // dtDuLieu.Rows.Add(list_Row[i].ToString());
                dtDuLieu2.Rows.Add(list_Row[i].ToString());
                dtDuLieu3.Rows.Add(list_Row[i].ToString());
            }
            // dtDuLieu.Rows.Add("Total");
            for (int i = 0; i < NoRow; i++)
            {
                for (int j = 0; j < NoCol; j++)
                {
                    //Dữ liệu 2
                    try
                    {
                        var dValue = from row in tblChartOK.AsEnumerable()
                                     where row.Field<string>(tblChartOK.Columns[0].ColumnName) == dtDuLieu2.Rows[i][0].ToString()
                                           && row.Field<string>(tblChartOK.Columns[1].ColumnName) == dtDuLieu2.Columns[j + 1].ColumnName
                                     select row.Field<double>(tblChartOK.Columns[3].ColumnName);
                        double Value_Ok = dValue.ToList()[0];
                        dtDuLieu2.Rows[i][j + 1] = Value_Ok/100;
                        dtDuLieu2.AcceptChanges();
                    }
                    catch (Exception)
                    {
                        dtDuLieu2.Rows[i][j + 1] = 0;
                        dtDuLieu2.AcceptChanges();
                    }
                    //Dữ liệu 3
                    try
                    {
                        var dValue = from row in tblChartOK.AsEnumerable()
                                     where row.Field<string>(tblChartOK.Columns[0].ColumnName) == dtDuLieu3.Rows[i][0].ToString()
                                           && row.Field<string>(tblChartOK.Columns[1].ColumnName) == dtDuLieu3.Columns[j + 1].ColumnName
                                     select row.Field<double>(tblChartOK.Columns[4].ColumnName);
                        double Value_Ok = dValue.ToList()[0];
                        dtDuLieu3.Rows[i][j + 1] = Value_Ok/100;
                        dtDuLieu3.AcceptChanges();
                    }
                    catch (Exception)
                    {
                        dtDuLieu3.Rows[i][j + 1] = 0;
                        dtDuLieu3.AcceptChanges();
                    }

                }
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; //fake key


            string SavePath = "";
            string DateTimeNow = DateTime.Now.Year.ToString()
            + string.Format("{0:00}", DateTime.Now.Month).ToString()
            + string.Format("{0:00}", DateTime.Now.Day).ToString() + "_"
            + string.Format("{0:00}", DateTime.Now.Hour).ToString()
            + string.Format("{0:00}", DateTime.Now.Minute).ToString()
            + string.Format("{0:00}", DateTime.Now.Second).ToString();

            using (ExcelPackage pck = new ExcelPackage())
            {
                //========================Định dạng sheet1=======================

                pck.Workbook.Properties.Author = "VNW0014732";
                pck.Workbook.Properties.Company = "FHS";
                pck.Workbook.Properties.Title = "Exported by FHS QuizTest";
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("匯總表");
                ExcelWorksheet ws2 = pck.Workbook.Worksheets.Add("堆疊直條圖");
                ExcelWorksheet ws3 = pck.Workbook.Worksheets.Add("百分比堆疊直條圖");
                //Định dạng toàn Sheet
                ws.Cells.Style.Font.Name = "Times New Roman";
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells.Style.Font.Size = 14;
                //Định dạng ô title
                ws.Cells[1, 1, 1, NoCol_ToTal + 1].Merge = true;
                ws.Row(1).Height = 30;
                ws.Cells["A1"].Value = "扁鋼胚品質統計軟體";
                ws.Cells["A1"].Style.Font.Size = 20;
                ws.Cells["A1"].Style.Font.Name = "DFKai-SB";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //Định dạng ô ngày tháng
                ws.Cells[2, 1, 2, NoCol_ToTal + 1].Merge = true;
                ws.Cells["A2"].Value = "日期: " + DateTime.Now.ToString().Split(' ')[0];
                ws.Cells["A2"].Style.Font.Bold = true;
                ws.Cells["A2"].Style.Font.Name = "DFKai-SB";
                ws.Cells["A2"].Style.Font.Size = 10;
                ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                //Định dạng ô title
                ws.Cells[3, 1, 3, NoCol_ToTal + 1].Merge = true;
                ws.Row(1).Height = 30;
                ws.Cells["A3"].Value = "統計表";
                ws.Cells["A3"].Style.Font.Size = 20;
                ws.Cells["A3"].Style.Font.Name = "DFKai-SB";
                ws.Cells["A3"].Style.Font.Bold = true;
                ws.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells["A4"].Value = "總共";
                ws.Cells["A4"].Style.Font.Name = "DFKai-SB";
                ws.Cells["A4"].Style.Font.Bold = true;

                //Đổ Background
                ws.Cells[4, 1, 4, NoCol_ToTal + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //ws.Cells[3, 1, 3, NoCol_ToTal + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells[4, 1, 4, NoCol_ToTal + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

                //Border
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Font.Size = 13;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Font.Name = "Times New Roman";
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.Font.Name = "DFKai-SB";
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[3, 1, NoRow_ToTal + 5, NoCol_ToTal + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //========================Định dạng sheet2=======================

                ws2.Cells.Style.Font.Name = "Times New Roman";
                ws2.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws2.Cells.Style.Font.Size = 14;
                //Định dạng ô title
                ws2.Cells[1, 1, 1, NoCol + 1].Merge = true;
                ws2.Row(1).Height = 30;
                ws2.Cells["A1"].Value = "扁鋼胚品質統計軟體";
                ws2.Cells["A1"].Style.Font.Size = 20;
                ws2.Cells["A1"].Style.Font.Name = "DFKai-SB";
                ws2.Cells["A1"].Style.Font.Bold = true;
                ws2.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //Định dạng ô ngày tháng
                ws2.Cells[2, 1, 2, NoCol + 1].Merge = true;
                ws2.Cells["A2"].Value = "日期: " + DateTime.Now.ToString().Split(' ')[0];
                ws2.Cells["A2"].Style.Font.Bold = true;
                ws2.Cells["A2"].Style.Font.Name = "DFKai-SB";
                ws2.Cells["A2"].Style.Font.Size = 10;
                ws2.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                //Định dạng ô title
                ws2.Cells[3, 1, 3, NoCol + 1].Merge = true;
                ws2.Row(1).Height = 30;
                ws2.Cells["A3"].Value = "統計表";
                ws2.Cells["A3"].Style.Font.Size = 20;
                ws2.Cells["A3"].Style.Font.Name = "DFKai-SB";
                ws2.Cells["A3"].Style.Font.Bold = true;
                ws2.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws2.Cells["A4"].Value = "總共";
                ws2.Cells["A4"].Style.Font.Name = "DFKai-SB";
                ws2.Cells["A4"].Style.Font.Bold = true;

                //Đổ Background
                ws2.Cells[3, 1, 3, NoCol + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws2.Cells[3, 1, 3, NoCol + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                //Border
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                ws2.Cells[6, 2, NoRow + 5, NoCol + 1].Style.Numberformat.Format = "##0.00%";

                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Size = 13;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Name = "Times New Roman";
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Name = "DFKai-SB";
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws2.Cells[3, 1, NoRow + 5, NoCol + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //========================Định dạng sheet3=======================

                ws3.Cells.Style.Font.Name = "Times New Roman";
                ws3.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws3.Cells.Style.Font.Size = 14;
                //Định dạng ô title
                ws3.Cells[1, 1, 1, NoCol + 1].Merge = true;
                ws3.Row(1).Height = 30;
                ws3.Cells["A1"].Value = "扁鋼胚品質統計軟體";
                ws3.Cells["A1"].Style.Font.Size = 20;
                ws3.Cells["A1"].Style.Font.Name = "DFKai-SB";
                ws3.Cells["A1"].Style.Font.Bold = true;
                ws3.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //Định dạng ô ngày tháng
                ws3.Cells[2, 1, 2, NoCol + 1].Merge = true;
                ws3.Cells["A2"].Value = "日期: " + DateTime.Now.ToString().Split(' ')[0];
                ws3.Cells["A2"].Style.Font.Bold = true;
                ws3.Cells["A2"].Style.Font.Name = "DFKai-SB";
                ws3.Cells["A2"].Style.Font.Size = 10;
                ws3.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                //Định dạng ô title
                ws3.Cells[3, 1, 3, NoCol + 1].Merge = true;
                ws3.Row(1).Height = 30;
                ws3.Cells["A3"].Value = "統計表";
                ws3.Cells["A4"].Value = "總共";
                ws3.Cells["A4"].Style.Font.Name = "DFKai-SB";
                ws3.Cells["A4"].Style.Font.Bold = true;
                ws3.Cells["A3"].Style.Font.Size = 20;
                ws3.Cells["A3"].Style.Font.Name = "DFKai-SB";
                ws3.Cells["A3"].Style.Font.Bold = true;
                ws3.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //3Đổ Background
                ws3.Cells[3, 1, 3, NoCol + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws3.Cells[3, 1, 3, NoCol + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                //Border
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Border.BorderAround(ExcelBorderStyle.Thick);

                ws3.Cells[6, 2, NoRow + 5, NoCol + 1].Style.Numberformat.Format = "##0.00%";

                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Size = 13;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Name = "Times New Roman";
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.Font.Name = "DFKai-SB";
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws3.Cells[3, 1, NoRow + 5, NoCol + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                //Thêm dữ liệu từ Grid vào Excel
                ws.Cells["A5"].LoadFromDataTable(dtDuLieu, true);
                for (int i = 0; i < NoCol_ToTal; i++)
                {
                    ws.Column(i + 2).Width = 20;
                    ws.Cells[4, i + 2].Value = SumCellVals(ws, i + 2, 6, NoRow_ToTal + 5);
                }
                //Thêm dữ liệu từ Grid vào Excel
                ws2.Cells["A5"].LoadFromDataTable(dtDuLieu2, true);
                for (int i = 0; i < NoCol; i++)
                {
                    //double SUM = dtCloned.AsEnumerable().Where(row => row.Field<object>(Column1_Name).ToString() == Value_Seach).Sum(row => row.Field<double>(Column2_Name));
                    double sum = dtDuLieu.AsEnumerable().Sum(x => x.Field<double>(dtDuLieu2.Columns[i + 1].ColumnName));
                    ws2.Column(i + 2).Width = 20;
                    ws2.Cells[4, i + 2].Value = sum;
                    //ws2.Cells[4, i + 2].Value = SumCellVals(ws2, i + 2, 6, NoRow + 5);
                }
                //Thêm dữ liệu từ Grid vào Excel
                ws3.Cells["A5"].LoadFromDataTable(dtDuLieu3, true);
                for (int i = 0; i < NoCol; i++)
                {
                    double sum = dtDuLieu.AsEnumerable().Sum(x => x.Field<double>(dtDuLieu3.Columns[i + 1].ColumnName));
                    ws3.Column(i + 2).Width = 20;
                    ws3.Cells[4, i + 2].Value = sum;
                }

                // Vẽ chart
                string nameChart = txbNameChart.Text;

                ExcelChart chart = ws.Drawings.AddChart(nameChart, eChartType.ColumnStacked);
                chart.XAxis.RemoveGridlines();
                chart.YAxis.RemoveGridlines();
                chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.ColumnChartStyle1);
                chart.Title.Text = nameChart;
                chart.SetSize(800, 400);

                chart.SetPosition(NoRow_ToTal + 6, 0, 0, 0); //chọn điểm để đặt đồ thị (điểm bắt đầu so với hàng,0,điểm bắt đầu so với côt, 0)
                chart.Legend.Remove();
                for (int i = 0; i < NoRow_ToTal; i++)
                {
                    //var Series = chart.Series.Add($"B{i + 3}:D{i + 3}", "B2:D2");
                    var Series = chart.Series.Add($"{ws.Cells[i + 6, 2, i + 6, NoCol_ToTal + 1]}", $"{ws.Cells[5, 2, 5, NoCol_ToTal + 1]}"); // add (vùng giá trị của series(0), gồm những cột nào)
                    chart.Series[i].Header = Convert.ToString(ws.Cells[i + 6, 1].Value);// tên các đường
                }
                //Format the legend
                chart.Legend.Add();
                chart.Legend.Border.Width = 0;
                chart.Legend.Font.Size = 10;
                //chart.Legend.Font.Bold = true;
                chart.Legend.Position = eLegendPosition.Right;


                string nameChart2 = txbNameChart.Text;
                ExcelChart chart2 = ws2.Drawings.AddChart(nameChart2, eChartType.ColumnStacked);
                chart2.XAxis.RemoveGridlines();
                chart2.YAxis.RemoveGridlines();
                chart2.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.ColumnChartStyle1);
                chart2.Title.Text = nameChart;
                chart2.SetSize(800, 400);
                if (SpinMaximum.Value!=0)
                {
                    chart2.YAxis.MaxValue = Convert.ToDouble(SpinMaximum.Value)/100;
                }
                

                chart2.SetPosition(NoRow + 6, 0, 0, 0); //chọn điểm để đặt đồ thị (điểm bắt đầu so với hàng,0,điểm bắt đầu so với côt, 0)
                chart2.Legend.Remove();
                for (int i = 0; i < NoRow; i++)
                {
                    //var Series = chart.Series.Add($"B{i + 3}:D{i + 3}", "B2:D2");
                    var Series = chart2.Series.Add($"{ws2.Cells[i + 6, 2, i + 6, NoCol + 1]}", $"{ws2.Cells[5, 2, 5, NoCol + 1]}"); // add (vùng giá trị của series(0), gồm những cột nào)
                    chart2.Series[i].Header = Convert.ToString(ws2.Cells[i + 6, 1].Value);// tên các đường
                }
                //Format the legend
                chart2.Legend.Add();
                chart2.Legend.Border.Width = 0;
                chart2.Legend.Font.Size = 10;
                //chart.Legend.Font.Bold = true;
                chart2.Legend.Position = eLegendPosition.Right;


                string nameChart3 = txbNameChart.Text;

                // var templateFile = new FileInfo("ok.crtx");
                // //var chart3 = (ExcelChart)ws3.Drawings.AddChartFromTemplate(templateFile, "areaChart");

                //// var areaChart = (ExcelAreaChart)ws.Drawings.AddChartFromTemplate(templateFile, "areaChart");
                // var areaChart = (ExcelAreaChart)ws.Drawings.AddChartFromTemplate(FileUtil.GetFileInfo("ok.crtx"), "areaChart");

                //var areaSerie = areaChart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
                //areaSerie.Header = "Order Value";
                //areaChart.SetPosition(51, 0, 10, 0);
                //areaChart.SetSize(1000, 300);
                //areaChart.Title.Text = "Area Chart";

                ExcelChart chart3 = ws3.Drawings.AddChart(nameChart3, eChartType.ColumnStacked100);
                chart3.XAxis.RemoveGridlines();
                chart3.YAxis.RemoveGridlines();
                chart3.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.ColumnChartStyle1);
                chart3.Title.Text = nameChart;
                chart3.Title.Font.SetFromFont(new System.Drawing.Font("DFKai-SB", 12));
                chart3.SetSize(800, 400);



               // chart.XAxis.MinValue = businessDayDate.AddDays(chartDayThreshold * -1).ToOADate();


                chart3.SetPosition(NoRow + 6, 0, 0, 0); //chọn điểm để đặt đồ thị (điểm bắt đầu so với hàng,0,điểm bắt đầu so với côt, 0)
                chart3.Legend.Remove();
                for (int i = 0; i < NoRow; i++)
                {
                    //var Series = chart.Series.Add($"B{i + 3}:D{i + 3}", "B2:D2");
                    var Series = chart3.Series.Add($"{ws3.Cells[i + 6, 2, i + 6, NoCol + 1]}", $"{ws3.Cells[5, 2, 5, NoCol + 1]}"); // add (vùng giá trị của series(0), gồm những cột nào)
                    chart3.Series[i].Header = Convert.ToString(ws3.Cells[i + 6, 1].Value);// tên các đường

                }
                //Format the legend
                chart3.Legend.Add();
                chart3.Legend.Border.Width = 0;
                chart3.Legend.Font.Size = 10;
                //chart.Legend.Font.Bold = true;
                chart3.Legend.Position = eLegendPosition.Right;


              

                FolderBrowserDialog Dialog = new FolderBrowserDialog();
                if (Dialog.ShowDialog() == DialogResult.OK)
                {
                    SavePath = Dialog.SelectedPath + "\\" + "Report_" + DateTimeNow + ".xlsx";
                }
                else
                {
                    return;
                }
                FileInfo excelFile = new FileInfo(SavePath);
                pck.SaveAs(excelFile);
            }
            Process.Start(SavePath);
        }
        private string SumCellVals(ExcelWorksheet ws, int colNum, int firstRow, int lastRow)
        {
            double runningTotal = 0.0;
            double currentVal;
            for (int i = firstRow; i <= lastRow; i++)
            {
                using (var sumCell = ws.Cells[i, colNum])
                {
                    currentVal = Convert.ToDouble(sumCell.Value);
                    runningTotal = runningTotal + currentVal;
                }
            }
            return runningTotal.ToString();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            string asda = pivotGridControl1.Prefilter.CriteriaString;
        }

        private void txtPath_EditValueChanged(object sender, EventArgs e)
        {

        }
    }

}

