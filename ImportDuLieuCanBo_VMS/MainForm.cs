using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Aspose.Cells;

namespace ImportDuLieuCanBo_VMS
{
    public partial class MainForm : Form
    {
        public string ConnectionString
        {
            get
            {
                var cnn = System.Configuration.ConfigurationManager.ConnectionStrings["VMS"].ConnectionString;

                if (cnn == "")
                    MessageBox.Show("Bạn chưa thiết lập chuỗi kết nối trong file Config.");

                return cnn;
            }
        }
        public string UploadPath
        {
            get
            {
                return System.Configuration.ConfigurationManager.AppSettings["IIS_Upload_Folder"];
            }
        }
        public SqlConnection Connection;
        private DBDataContext db;
        private DataTable dtDonVi = new DataTable();


        public MainForm()
        {
            InitializeComponent();
        }

        #region Procedure Process

        private void LoadImportList()
        {
            try
            {
                var selectedFolder = new DirectoryInfo(txtFolder.Text);

                if (!selectedFolder.Exists)
                {
                    MessageBox.Show("Bạn chưa chọn thư mục hoặc thư mục không tồn tại.");
                    return;
                }

                var folderList = selectedFolder.GetDirectories();
                listViewFolder.Items.Clear();
                foreach (var folder in folderList)
                {
                    var fileCheck = folder.GetFiles("*.xls");
                    if (fileCheck.Length > 0)
                    {
                        var listItem = new ListViewItem(folder.Name);
                        listItem.Checked = true;

                        listItem.ImageKey = "status-offline.png";

                        listItem.SubItems.Add("");

                        listViewFolder.Items.Add(listItem);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void LogMessage(string message)
        {
            var resultMessage = string.Format("{0:HH:mm}: {1}\n", DateTime.Now, message);

            logWindow.AppendText(resultMessage);
        }

        private void StartImport(List<ListViewItem> items)
        {
            txtFolder.Enabled =
                btnBrowseFolder.Enabled =
                btnLoadList.Enabled =
                btnImportAll.Enabled =
                btnImportSelected.Enabled =
                lnkSelectAll.Enabled =
                lnkSelectNone.Enabled = false;

            btnStop.Enabled = true;

            var err = "";

            dtDonVi = SqlGet("SELECT * FROm DonVi", new SqlParameter[] { }, out err);

            importWorker.RunWorkerAsync(items);
        }

        #region Sql Helper
        private DataTable SqlGet(string query, SqlParameter[] parameters, out string errMessage)
        {
            DataTable table = new DataTable();

            errMessage = "";

            try
            {
                if (Connection == null)
                {
                    Connection = new SqlConnection(ConnectionString);
                    Connection.Open();
                    using (SqlDataAdapter da = new SqlDataAdapter(query, Connection))
                    {
                        da.Fill(table);

                        return table;
                    }
                }
                else
                {
                    SqlDataAdapter da = new SqlDataAdapter(query, Connection);

                    da.Fill(table);

                    return table;
                }

            }
            catch (SqlException ex)
            {
                errMessage += ex.ToString();
                return null;
            }

            return table;
        }

        private int SqlExecute(string query, SqlParameter[] parameters, out string errMessage)
        {
            int rowApply = 0;
            errMessage = "";
            try
            {
                using (var command = Connection.CreateCommand())
                {
                    command.CommandText = query;
                    command.CommandType = CommandType.Text;
                    command.Parameters.AddRange(parameters);

                    rowApply = command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
            }
            return rowApply;
        }
        #endregion

        private int GetDanhMuc_ID(string loaiDanhMuc, string tenDanhMuc, out string err)
        {
            var id = 0;

            var sql = string.Format("SELECT {0}_ID FROM DM_{0} WHERE Ten{0} LIKE N'{1}_' OR Ten{0} LIKE N'{1}'", loaiDanhMuc, tenDanhMuc);

            var data = SqlGet(sql, new SqlParameter[] { }, out err);

            if (err == "")
            {
                if (data.Rows.Count > 0)
                {
                    return int.Parse(data.Rows[0][loaiDanhMuc + "_ID"] + "");
                }

                if (SqlExecute(string.Format("INSERT INTO DM_{0}(Ten{0}) VALUES (N'{1}')", loaiDanhMuc, tenDanhMuc), new SqlParameter[] { }, out err) > 0)
                {
                    return int.Parse(SqlGet(string.Format("SELECT MAX({0}_ID) FROM DM_{0}", loaiDanhMuc), new SqlParameter[] { }, out err).Rows[0][0] + "");
                }

                err += "; data = " + tenDanhMuc;
            }

            return id;
        }
        private int GetDanhMucFixed_ID(string loaiDanhMuc, string tenDanhMuc, out string err)
        {
            var id = 0;

            var sql = string.Format("SELECT {0}_ID FROM FixedDM_{0} WHERE Ten{0} LIKE N'{1}'", loaiDanhMuc,tenDanhMuc);

            err = "";
            var data = SqlGet(sql, new SqlParameter[] { }, out err);
            

            if (err == "")
            {
                if (data.Rows.Count > 0)
                {
                    return int.Parse(data.Rows[0][loaiDanhMuc + "_ID"] + "");
                }
            }

            return id;
        }

        private int GetDonVi_ID(string[] donVis)
        {
            int id = 0;

            var listSort = new List<string>();

            for (int i = donVis.Length - 1; i >= 0; i--)
            {
                listSort.Add(donVis[i].Trim());
            }

            var level = 0;
            var id_parent = -1;

            foreach (string tenDonVi in listSort)
            {
                DataRow r = dtDonVi.AsEnumerable().FirstOrDefault(p => p.Field<string>("TenDonVi") == tenDonVi && (level == 0 || p.Field<int?>("id_CapTren") == id_parent));

                if (r != null)
                {
                    id = r.Field<int>("DonVi_ID");

                    level++;
                    id_parent = r.Field<int>("DonVi_ID");
                }
                else
                {
                    break;
                }
            }

            if (level != listSort.Count)
            {
                id = 0;
            }

            return id;
        }
        #endregion

        #region events
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (ConnectionString == "")
            {
                this.Close();
            }

            if (Connection == null)
            {
                try
                {
                    Connection = new SqlConnection(ConnectionString);
                    Connection.Open();
                    db = new DBDataContext(ConnectionString);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi khi try cập CSDL: " + ex.Message);
                    this.Close();
                }
            }




            txtFolder.Text = "F:/File";
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Connection.State == ConnectionState.Open)
                Connection.Close();
        }

        private void btnLoadList_Click(object sender, EventArgs e)
        {
            LoadImportList();
        }

        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.SelectedPath = txtFolder.Text;
            if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFolder.Text = folderBrowserDialog.SelectedPath;
                LoadImportList();
            }
        }

        private void lnkSelectAll_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            foreach (ListViewItem item in listViewFolder.Items)
            {
                item.Checked = true;
            }
        }

        private void lnkSelectNone_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            foreach (ListViewItem item in listViewFolder.Items)
            {
                item.Checked = false;
            }
        }

        private void listViewFolder_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void logWindow_TextChanged(object sender, EventArgs e)
        {
            logWindow.ScrollToCaret();
        }

        private void btnImportSelected_Click(object sender, EventArgs e)
        {
            var importList = new List<ListViewItem>();

            foreach (ListViewItem item in listViewFolder.CheckedItems)
            {
                importList.Add(item);
            }

            StartImport(importList);
        }

        private void btnImportAll_Click(object sender, EventArgs e)
        {
            var importList = new List<ListViewItem>();

            foreach (ListViewItem item in listViewFolder.Items)
            {
                importList.Add(item);
            }

            StartImport(importList);
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            importWorker.CancelAsync();
        }

        private void importWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string[] obj = e.UserState as string[];

            var itemText = obj[0] as string;
            var item = listViewFolder.FindItemWithText(itemText);
            if (item != null)
            {
                item.ImageKey = obj[1] as string;

                if (item.ImageKey == "status.png")
                {
                    item.SubItems[1].Text = "OK";
                }
                else if (item.ImageKey == "status-busy.png")
                {
                    item.SubItems[1].Text = "Lỗi";
                }
            }

            var msg = obj[2] as string;

            LogMessage(msg);
        }

        private void importWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LogMessage("Dừng import");
            txtFolder.Enabled =
                btnBrowseFolder.Enabled =
                btnLoadList.Enabled =
                btnImportAll.Enabled =
                btnImportSelected.Enabled =
                lnkSelectAll.Enabled =
                lnkSelectNone.Enabled = true;

            btnStop.Enabled = false;
        }
        public bool checkDigit(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                if (char.IsDigit(s[i]) == true)
                    return true;
            }
            return false;
        }

        private void importWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<ListViewItem> items = e.Argument as List<ListViewItem>;

            var err = "";

            foreach (var item in items)
            {
                if (importWorker.CancellationPending)
                {
                    break;
                }

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Bắt đầu Import thư mục: \"{0}\"", item.Text) });

                var directory = new DirectoryInfo(txtFolder.Text + "\\" + item.Text);

                #region Chuan bi file Import
                var excelFile = directory.GetFiles("*.xls")[0].FullName;
                Workbook workbook = new Workbook();
                Worksheet sheet;
                Cells cells;

                try
                {
                    workbook.Open(excelFile);
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi mở tệp: \"{0}\"", ex.Message) });
                    continue;
                }
                #endregion

                #region Import ThongTinCoBan

                #region buildDuLieuImport
                // Nếu có cán bộ rồi thì update thông tin giao tiếp
                // Nếu chưa có thì thêm mới
                // Không lấy ID cán bộ theo file excel mà dựa vào cột mã cán bộ. Từ đó lấy ra ID để lưu vào các bảng có quan hệ 1-n, trong trường hợp này là quá trình công tác
                // 
                sheet = workbook.Worksheets["ThongTinChung"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lấy dữ liệu ThongTinChung") });
                var canbo = new Entities.CanBo();
                //for (int i = 4; i < sheet.Cells.MaxRow; i++)
                //{
                //    int j = 0;
                    //if (cells[i, j].Value + "" == "")
                    //{
                    //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lấy dữ liệu ThongTinChung: Chưa nhập ID") });
                    //    continue;
                    //}
                    if (cells["A6"].Value + "" == "")
                    {
                        importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lấy dữ liệu ThongTinChung: Chưa nhập ID") });
                        continue;
                    }
                    try
                    {
                        //canbo.CanBo_ID = cells[i, j].StringValue;

                        //canbo.MaCanBo = cells[i, j + 1].StringValue;
                        //canbo.MaThe = cells[i, j + 2].StringValue;
                        //canbo.ChiMucHoSoGoc = cells[i, j + 3].StringValue;
                        //canbo.HoDem = cells[i, j + 4].StringValue;
                        //canbo.Ten = cells[i, j + 5].StringValue;
                        //canbo.ChucVuKiemNhiem = cells[i, j + 6].StringValue;
                        //canbo.DienThoaiCQ = cells[i, j + 7].StringValue;
                        //canbo.DTDD = cells[i, j + 8].StringValue;
                        //canbo.Email = cells[i, j + 9].StringValue;
                        //canbo.SoTaiKhoan = cells[i, j + 10].StringValue;
                        //canbo.TenNganHang = cells[i, j + 11].StringValue;
                        //canbo.IDNganHang = GetDanhMuc_ID("NganHang", canbo.TenNganHang, out err);
                        //canbo.SoCMND = cells[i, j + 12].StringValue;
                        //canbo.NgayCapCMND = cells[i, j + 13].StringValue;
                        //canbo.NoiCapCMND = cells[i, j + 14].StringValue;
                        //canbo.HoVaTenKhaiSinh = canbo.HoDem + canbo.Ten;
                        //canbo.GioiTinh = cells[i, j + 15].StringValue.ToLower() == "1";
                        //canbo.NgaySinh = cells[i, j + 16].StringValue;
                        //canbo.NoiSinh = cells[i, j + 17].StringValue;
                        //canbo.QueQuan = cells[i, j + 18].StringValue;
                        //canbo.ThuongTru = cells[i, j + 19].StringValue;
                        //canbo.TamTru = cells[i, j + 20].StringValue;
                        //canbo.TenDanToc = cells[i, j + 21].StringValue;
                        //canbo.ID_DanToc = GetDanhMuc_ID("DanToc", canbo.TenDanToc, out err);
                        //canbo.TenTonGiao = cells[i, j + 22].StringValue;
                        //canbo.ID_TonGiao = GetDanhMuc_ID("TonGiao", canbo.TenTonGiao, out err);
                        //canbo.NgayVaoCongTy = cells[i, j + 23].StringValue;
                        //canbo.NgayVaoNganh = cells[i, j + 25].StringValue;
                        //canbo.NgayThoiViec = cells[i, j + 24].StringValue;
                        //canbo.NgayVaoDang = cells[i, j + 26].StringValue;
                        //canbo.TrinhDoVanHoa = cells[i, j + 27].StringValue;
                        //canbo.TrinhDoDaoTaoChiTiet = cells[i, j + 27].StringValue;
                        //canbo.TrinhDoTinHoc = cells[i, j +28].StringValue;
                        //canbo.TinhTrangHonNhan = cells[i, j + 29].StringValue;



                        canbo.CanBo_ID = cells["A6"].StringValue;
                        canbo.MaCanBo = cells["C6"].StringValue;
                        canbo.HoDem = cells["D6"].StringValue;
                        canbo.Ten = cells["E6"].StringValue;
                        canbo.GioiTinh = cells["F6"].StringValue.ToLower() == "Nam";
                        canbo.NgaySinh = cells["G6"].StringValue;
                        canbo.NoiSinh = cells["J6"].StringValue;
                        canbo.QueQuan = cells["K6"].StringValue;
                        canbo.ThuongTru = cells["L6"].StringValue;
                        canbo.TamTru = cells["M6"].StringValue;
                        canbo.DTDD = cells["N6"].StringValue;
                        canbo.TenDanToc = cells["O6"].StringValue;
                        canbo.ID_DanToc = GetDanhMuc_ID("DanToc", canbo.TenDanToc, out err);
                        canbo.TenTonGiao = cells["P6"].StringValue;
                        canbo.ID_TonGiao = GetDanhMuc_ID("TonGiao", canbo.TenTonGiao, out err);
                        canbo.NgayVaoCongTy = cells["Q6"].StringValue;
                        canbo.SoCMND = cells["R6"].StringValue;
                        canbo.NgayCapCMND = cells["S6"].StringValue;
                        canbo.NoiCapCMND = cells["T6"].StringValue;
                        canbo.NgayVaoDang = cells["U6"].StringValue;
                        canbo.NgayNhapNgu = cells["V6"].StringValue;
                        canbo.NgayXuatNgu = cells["W6"].StringValue;
                        canbo.NhomMau = cells["AA6"].StringValue;






                    }
                    //catch (Exception ex)
                    //{
                    //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lấy dữ liệu ThongTinChung: \"{0}\"", ex.Message) });
                    //    continue;
                    //}
                    finally { }

                    #endregion
                    #region Thong Tin Giao Tiep

                    importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu thông tin giao tiếp") });
                    // ??? CHỗ này để lọc lấy ra Thông tin cán bộ (Đang dùng LInq)
                    // Thay vì lấy theo ID thì phải lấy theo mã cán bộ. Biến đối tương tự cho phần cán bộ - quá trình công tác
                    var ttCanBo = db.CanBos.FirstOrDefault(p => p.MaCanBo == canbo.MaCanBo);
                    try
                    {
                        //ttCanBo = new CanBo();
                        ////ttCanBo.CanBo_ID = int.Parse(canbo.CanBo_ID);
                        //ttCanBo.MaCanBo = canbo.MaCanBo;
                        //ttCanBo.MaThe = canbo.MaThe;
                        //ttCanBo.ChiMucHoSoGoc = canbo.ChiMucHoSoGoc;
                        //ttCanBo.HoDem = canbo.HoDem;
                        //ttCanBo.Ten = canbo.Ten;
                        //ttCanBo.ChucVuKiemNhiem = canbo.ChucVuKiemNhiem;
                        //ttCanBo.DienThoaiCQ = canbo.DienThoaiCQ;
                        //ttCanBo.DienThoaiDD = canbo.DTDD;
                        //ttCanBo.Email = canbo.Email;
                        //ttCanBo.SoTaiKhoan = canbo.SoTaiKhoan;
                        //ttCanBo.ID_NganHang = canbo.IDNganHang;
                        //ttCanBo.SoCMND = canbo.SoCMND;
                        //ttCanBo.NgayCapCMND = canbo.NgayCapCMND;
                        //ttCanBo.NoiCapCMND = canbo.NoiCapCMND;
                        //ttCanBo.HoTenKhaiSinh = canbo.HoVaTenKhaiSinh;
                        //ttCanBo.GioiTinh = canbo.GioiTinh;
                        //ttCanBo.NgaySinh = canbo.NgaySinh;
                        //ttCanBo.NoiSinh = canbo.NoiSinh;
                        //ttCanBo.QueQuan_DiaChi = canbo.QueQuan;
                        //ttCanBo.ThuongTru = canbo.ThuongTru;
                        //ttCanBo.TamTru = canbo.TamTru;
                        //ttCanBo.ID_DanToc = (short)canbo.ID_DanToc;
                        //ttCanBo.ID_TonGiao = (short)canbo.ID_TonGiao;

                        //ttCanBo.NgayVaoCongTy = canbo.NgayVaoCongTy;
                        //ttCanBo.NgayVaoNganh = canbo.NgayVaoNganh;
                        //ttCanBo.NgayThoiViec = canbo.NgayThoiViec;
                        //ttCanBo.TrinhDoVanHoa = canbo.TrinhDoVanHoa;
                        //ttCanBo.TrinhDoDaoTaoChiTiet = canbo.TrinhDoDaoTaoChiTiet;
                        //ttCanBo.TrinhDoTinHoc = canbo.TrinhDoTinHoc;
                        ////ttCanBo.ID_TinhTrangHonNhan = (short)int.Parse(canbo.TinhTrangHonNhan);
                        //db.CanBos.InsertOnSubmit(ttCanBo);
                        //db.SubmitChanges();



                         ttCanBo = new CanBo();
                        ttCanBo.MaCanBo = canbo.MaCanBo;
                        ttCanBo.HoDem = canbo.HoDem;
                        ttCanBo.Ten = canbo.Ten;
                        ttCanBo.GioiTinh = canbo.GioiTinh;
                        ttCanBo.NgaySinh = canbo.NgaySinh;
                        ttCanBo.NoiSinh = canbo.NoiSinh;
                        ttCanBo.ID_DanToc = (short)canbo.ID_DanToc;
                        ttCanBo.ID_TonGiao = (short)canbo.ID_TonGiao;
                        ttCanBo.QueQuan_DiaChi = canbo.QueQuan;
                        ttCanBo.ThuongTru = canbo.ThuongTru;
                        ttCanBo.TamTru = canbo.TamTru;
                        ttCanBo.NgayVaoCongTy = canbo.NgayVaoCongTy;
                        ttCanBo.NhomMau = canbo.NhomMau;
                         db.CanBos.InsertOnSubmit(ttCanBo);
                        db.SubmitChanges();

                        // ??? Phải lấy được CanBo_ID
                        // var ttCanBMax = db.CanBos.FirstOrDefault(p => p.MaCanBo == canbo.MaCanBo);
                    }
                    //catch (Exception ex)
                    //{
                    //    if (db.Transaction != null) db.Transaction.Rollback();
                    //    err = ex.Message;
                    //}
                    finally { }

               // }
                var ttCanBoMax = db.CanBos.FirstOrDefault(p => p.MaCanBo == canbo.MaCanBo);
                if (err != "")
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi cập nhật thông tin giao tiếp: \"{0}\"", err) });
                    err = "";
                    continue;
                }
                #endregion
                //#region CanBoDang
                //if (importWorker.CancellationPending)
                //{
                //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                //    break;
                //}

                //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu thông tin đảng") });

                //try
                //{
                //    var bc = db.CanBoDangs.FirstOrDefault(p => p.CanBo_ID == int.Parse(canbo.CanBo_ID));
                //    if (bc == null)
                //    {
                //        bc = new CanBoDang { CanBo_ID = int.Parse(canbo.CanBo_ID) };

                //        db.CanBoDangs.InsertOnSubmit(bc);
                //    }

                //    bc.NgayVaoDang = canbo.NgayVaoDang;

                //    db.SubmitChanges();
                //}
                //catch (Exception ex)
                //{
                //    if (db.Transaction != null) db.Transaction.Rollback();
                //    err = ex.Message;
                //}

                //if (err != "")
                //{
                //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi cập nhật thông tin Đảng: \"{0}\"", err) });
                //    err = "";
                //    continue;
                //}
                //#endregion
                //#region CanBoThamGiaQuanDoi
                //if (importWorker.CancellationPending)
                //{
                //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                //    break;
                //}

                //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu thông tin tham gia quân đội") });

                //try
                //{
                //    var bc = db.CanBoThamGiaQuanDois.FirstOrDefault(p => p.CanBo_ID == int.Parse(canbo.CanBo_ID));
                //    if (bc == null)
                //    {
                //        bc = new CanBoThamGiaQuanDoi { CanBo_ID = int.Parse(canbo.CanBo_ID) };

                //        db.CanBoThamGiaQuanDois.InsertOnSubmit(bc);
                //    }

                //    bc.NgayNhapNgu = canbo.NgayNhapNgu;
                //    bc.NgayXuatNgu = canbo.NgayXuatNgu;

                //    db.SubmitChanges();
                //}
                //catch (Exception ex)
                //{
                //    if (db.Transaction != null) db.Transaction.Rollback();
                //    err = ex.Message;
                //}

                //if (err != "")
                //{
                //    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi cập nhật thông tin tham gia quân đội: \"{0}\"", err) });
                //    err = "";
                //    continue;
                //}
                //#endregion
                #endregion
                #region Import QuaTrinhCongTac
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["CongTac"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu CongTac") });

                var startRow = 6;
                try
                {
                    var bc = new CanBo_QuaTrinhCongTac();
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "" || (cells["K" + startRow].StringValue + "" != "" && cells["A" + startRow].StringValue == cells["K" + startRow].StringValue))
                        {
                            startRow++;
                            continue;
                        }
                        
                        //??? ĐÚng là cái này cũng khó cho em. Khó về mặt kỹ thuật (Có cả linq)
                        // Nhưng mấu chốt, anh muốn em nhìn ra cái logic của vấn đề, qua đó tìm ra cái mình chưa hiểu
                        // Nếu có quá trình công tác ID thì lọc theo quá trình công tác
                        if (cells["A" + startRow].StringValue + "" != "" && checkDigit(cells["A" + startRow].StringValue.ToString()))
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_QuaTrinhCongTacs.FirstOrDefault(p => p.CanBo_QuaTrinhCongTac_ID == id);
                        }
                        else
                        {
                            // Nếu không thì lọc theo CanBoID ở trên
                            bc = db.CanBo_QuaTrinhCongTacs.FirstOrDefault(p => p.TuNgay == cells["F" + startRow].StringValue && p.ID_CanBo == int.Parse(canbo.CanBo_ID)) ??
                                 new CanBo_QuaTrinhCongTac { CanBo_QuaTrinhCongTac_ID = 0, ID_CanBo = int.Parse(canbo.CanBo_ID) };
                        }

                        bc.ID_LoaiQTCongTac = (byte)GetDanhMucFixed_ID("LoaiQTCongTac", cells["C" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu CongTac: \"{0}\"", err) });
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.SoQuyetDinh = cells["D" + startRow].StringValue;
                        bc.NgayKy = cells["E" + startRow].StringValue;
                        bc.TuNgay = cells["F" + startRow].StringValue;

                        var donvi_temp = cells["G" + startRow].StringValue;
                        var id_donvi = GetDonVi_ID(donvi_temp.Split(",".ToCharArray()));
                        if (id_donvi != 0)
                        {
                            bc.ID_DonVi = id_donvi;
                        }
                        else
                        {
                            bc.ID_DonVi = (int?)null;
                            bc.TenDonVi = donvi_temp;
                        }

                        bc.ID_ChucDanh = (short)GetDanhMuc_ID("ChucDanh", cells["H" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu CongTac: \"{0}\"", err) });
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.ChucDanh = cells["H" + startRow].StringValue;
                        bc.ID_ChucVu = (short)GetDanhMuc_ID("ChucVu", cells["I" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu CongTac: \"{0}\"", err) });
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.ChucVu = cells["I" + startRow].StringValue;
                        if (cells["K" + startRow].StringValue.Trim() + "" != "" && cells["A" + startRow].StringValue != cells["K" + startRow].StringValue)
                        {
                            bc.TenVanBan = cells["K" + startRow].StringValue +
                                           (cells["K" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");
                        }

                        if (bc.CanBo_QuaTrinhCongTac_ID == 0)
                        {
                            db.CanBo_QuaTrinhCongTacs.InsertOnSubmit(bc);
                        }

                        try
                        {
                            db.CanBo_QuaTrinhCongTacs.InsertOnSubmit(bc);
                            db.SubmitChanges();
                            
                        }
                        catch
                        {
                            db.Refresh(RefreshMode.OverwriteCurrentValues, bc);
                        }


                        // process file Upload

                        if (bc.TenVanBan + "" != "")
                        {
                            var fPath = Path.Combine(UploadPath, "QuaTrinhCongTacCanBo",
                                bc.CanBo_QuaTrinhCongTac_ID + "");

                            if (!(new FileInfo(fPath).Exists))
                            {
                                var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanBan));

                                if (fUpload.Exists)
                                {
                                    fUpload.CopyTo(fPath, true);
                                }
                            }
                        }
                        // process file Upload

                        startRow++;
                    }
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu CongTac: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }
                return;
                #endregion
                #region Import BangCap
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["BangCap"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu BangCap") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        err = "";
                        if (cells["C" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var bc = new CanBo_BangCap();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_BangCaps.FirstOrDefault(p => p.CanBo_BangCap_ID == id);
                        }
                        else
                        {
                            var cn = GetDanhMuc_ID("ChuyenNganh", cells["C" + startRow].StringValue.Trim(), out err);

                            bc = db.CanBo_BangCaps.FirstOrDefault(p => p.ID_CanBo == int.Parse(canbo.CanBo_ID) && p.ID_ChuyenNganh == cn) ??
                                 new CanBo_BangCap { CanBo_BangCap_ID = 0, ID_CanBo = int.Parse(canbo.CanBo_ID) };
                        }

                        bc.ID_ChuyenNganh = GetDanhMuc_ID("ChuyenNganh", cells["C" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu BangCap: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.ID_CapDaoTao = (short)GetDanhMuc_ID("CapDaoTao", cells["D" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu BangCap: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.ID_HeDaoTao = (short)GetDanhMuc_ID("HeDaoTao", cells["E" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu BangCap: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.NoiCapBang = cells["F" + startRow].StringValue;
                        bc.ID_VanBang = (short)GetDanhMuc_ID("VanBang", cells["G" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu BangCap: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }

                        bc.NamTotNghiep = cells["H" + startRow].Value == null ? (int?)null : cells["H" + startRow].IntValue;

                        if (err == "")
                        {
                            if (bc.CanBo_BangCap_ID == 0)
                            {

                                db.CanBo_BangCaps.InsertOnSubmit(bc);
                            }

                            try
                            {
                                db.SubmitChanges();
                            }
                            catch (Exception)
                            {
                                if (db.Transaction != null) db.Transaction.Rollback();
                            }

                            // process file Upload
                            if (cells["I" + startRow].StringValue != bc.CanBo_BangCap_ID + "")
                            {

                                bc.TenVanBan = cells["I" + startRow].StringValue +
                                               (cells["I" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");
                                if (bc.TenVanBan != "")
                                {
                                    var fPath = Path.Combine(UploadPath, "BangCap",
                                        bc.CanBo_BangCap_ID + "");

                                    if (!(new FileInfo(fPath).Exists))
                                    {
                                        var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanBan));

                                        if (fUpload.Exists)
                                        {
                                            fUpload.CopyTo(
                                                Path.Combine(UploadPath, "BangCap", bc.CanBo_BangCap_ID + ""),
                                                true);
                                        }
                                    }

                                }


                            }
                            // process file Upload
                        }
                        startRow++;
                    }

                    //db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu BangCap: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }

                #endregion
                #region Import ChungChi
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["ChungChi"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu ChungChi") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var bc = new CanBo_ChungChi();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_ChungChis.FirstOrDefault(p => p.CanBo_ChungChi_ID == id);
                        }
                        else
                        {
                            bc = db.CanBo_ChungChis.FirstOrDefault(p => p.ID_CanBo == int.Parse(canbo.CanBo_ID) && p.ChungChi == cells["C" + startRow].StringValue) ??
                                 new CanBo_ChungChi { CanBo_ChungChi_ID = 0, ID_CanBo = int.Parse(canbo.CanBo_ID) };
                        }

                        bc.ChungChi = cells["C" + startRow].StringValue;
                        bc.ID_XepLoai = (short)GetDanhMuc_ID("XepLoai", cells["D" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu ChungChi: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        if (cells["E" + startRow].StringValue == "")
                        {
                            bc.TuNgay = cells["G" + startRow].StringValue;
                        }
                        else
                        {
                            bc.TuNgay = cells["E" + startRow].StringValue;
                        }

                        bc.DenNgay = cells["F" + startRow].StringValue;
                        bc.NgayCap = cells["G" + startRow].StringValue;
                        bc.NoiCap = cells["H" + startRow].StringValue;

                        if (bc.CanBo_ChungChi_ID == 0)
                        {
                            db.CanBo_ChungChis.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        // process file Upload
                        if (cells["O" + startRow].StringValue != bc.CanBo_ChungChi_ID + "")
                        {
                            bc.TenVanBan = cells["I" + startRow].StringValue + (cells["I" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");
                            if (bc.TenVanBan != "")
                            {
                                var fPath = Path.Combine(UploadPath, "ChungChi",
                                    bc.CanBo_ChungChi_ID + "");

                                if (!(new FileInfo(fPath).Exists))
                                {
                                    var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanBan));

                                    if (fUpload.Exists)
                                    {
                                        fUpload.CopyTo(Path.Combine(UploadPath, "ChungChi", bc.CanBo_ChungChi_ID + ""),
                                            true);
                                    }
                                }
                            }
                        }
                        // process file Upload

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu ChungChi: \"{0}\"", ex.Message) });
                    err = "";
                    db.Refresh(RefreshMode.OverwriteCurrentValues, db);
                    continue;
                }

                #endregion
                #region Import GiaDinh
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["GiaDinh"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu ThanNhan") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "")
                        {
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }

                        var bc = new CanBo_ThanNhan();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_ThanNhans.FirstOrDefault(p => p.CanBo_ThanNhan_ID == id);
                        }
                        else
                        {
                            bc = db.CanBo_ThanNhans.FirstOrDefault(p => p.ID_CanBo == int.Parse(canbo.CanBo_ID) && p.HoTen == cells["C" + startRow].StringValue)
                                ?? new CanBo_ThanNhan() { CanBo_ThanNhan_ID = 0, ID_CanBo = int.Parse(canbo.CanBo_ID) };
                        }

                        bc.HoTen = cells["C" + startRow].StringValue;
                        bc.ID_QuanHe = (short)GetDanhMuc_ID("QuanHe", cells["D" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu ThanNhan: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.NamSinh = cells["E" + startRow].StringValue;
                        bc.QueQuan = cells["F" + startRow].StringValue;
                        bc.NoiOHienNay = cells["G" + startRow].StringValue;
                        bc.ThongTinCongTac = cells["H" + startRow].StringValue;

                        if (bc.CanBo_ThanNhan_ID == 0)
                        {
                            db.CanBo_ThanNhans.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu ThanNhan: \"{0}\"", ex.Message) });
                    err = "";
                    db.Refresh(RefreshMode.OverwriteCurrentValues, db);
                    continue;
                }

                #endregion
                #region Import HopDong
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["HopDong"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu HopDong") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "" || cells["O" + startRow].StringValue + "" == "" || cells["O" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var bc = new CanBo_HopDongLaoDong();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_HopDongLaoDongs.FirstOrDefault(p => p.HopDongID == id);
                        }
                        else
                        {
                            bc = db.CanBo_HopDongLaoDongs.FirstOrDefault(p => p.CanBo_ID == int.Parse(canbo.CanBo_ID) && p.TuNgay == cells["E" + startRow].StringValue)
                                ?? new CanBo_HopDongLaoDong() { HopDongID = 0, CanBo_ID = int.Parse(canbo.CanBo_ID) };
                        }

                        bc.SoHopDong = cells["C" + startRow].StringValue;
                        bc.ID_LoaiHopDong = (short)GetDanhMuc_ID("LoaiHopDong", cells["D" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu HopDong: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.TuNgay = cells["E" + startRow].StringValue;
                        bc.DenNgay = cells["F" + startRow].StringValue;
                        bc.NgayKyHopDong = cells["G" + startRow].StringValue;
                        bc.NgayHopDongCoHieuLuc = cells["H" + startRow].StringValue;
                        bc.CongViecPhaiLam = cells["I" + startRow].StringValue;
                        bc.DiaDiemLamViec = cells["J" + startRow].StringValue;
                        bc.MucLuong = cells["K" + startRow].StringValue;
                        bc.CacPhuCap = cells["L" + startRow].StringValue;
                        bc.NguoiKy = cells["M" + startRow].StringValue;

                        if (bc.HopDongID == 0)
                        {
                            db.CanBo_HopDongLaoDongs.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        if (cells["O" + startRow].StringValue != bc.HopDongID + "")
                        {
                            bc.TenVanBan = cells["O" + startRow].StringValue +
                                           (cells["O" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");

                            // process file Upload
                            if (bc.TenVanBan != "")
                            {
                                var fPath = Path.Combine(UploadPath, "HopDongLaoDong",
                                    bc.HopDongID + "");

                                if (!(new FileInfo(fPath).Exists))
                                {
                                    var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanBan));

                                    if (fUpload.Exists)
                                    {
                                        fUpload.CopyTo(Path.Combine(UploadPath, "HopDongLaoDong", bc.HopDongID + ""),
                                            true);
                                    }
                                }
                            }
                            // process file Upload
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu HopDong: \"{0}\"", ex.Message) });
                    err = "";
                    continue;
                }

                #endregion
                #region Import DanhHieu
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["DanhHieu"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu DanhHieu") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue + "" == "" || cells["I" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var bc = new CanBo_DanhHieu();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_DanhHieus.FirstOrDefault(p => p.CanBo_DanhHieu_ID == id);
                        }
                        else
                        {
                            bc.CanBo_DanhHieu_ID = 0;
                            bc.ID_CanBo = int.Parse(canbo.CanBo_ID);
                        }

                        bc.VaoNam = cells["C" + startRow].StringValue;
                        bc.ID_DanhHieu = (short)GetDanhMuc_ID("DanhHieu", cells["D" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu DanhHieu: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.SoQuyetDinh = cells["E" + startRow].StringValue;
                        bc.NgayRaQuyetDinh = cells["F" + startRow].StringValue;
                        bc.ID_NoiBanHanh = (short)GetDanhMuc_ID("NoiBanHanh", cells["G" + startRow].StringValue.Trim(), out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu DanhHieu: \"{0}\"", err) });
                            startRow++;
                            continue;
                        }

                        if (bc.CanBo_DanhHieu_ID == 0)
                        {
                            db.CanBo_DanhHieus.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu DanhHieu: \"{0}\"", ex.Message) });
                    if (db.Transaction != null) db.Transaction.Rollback();
                    err = "";
                    continue;
                }

                #endregion
                #region Import KhenThuong
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["KhenThuong"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu KhenThuong") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "")
                        {
                            startRow++;
                            continue;
                        }

                        var bc = new CanBo_KhenThuong();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_KhenThuongs.FirstOrDefault(p => p.CanBo_KhenThuong_ID == id);
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Sửa dữ liệu KhenThuong: " + id) });
                        }
                        else
                        {
                            bc.CanBo_KhenThuong_ID = 0;
                            bc.ID_CanBo = int.Parse(canbo.CanBo_ID);
                        }

                        bc.Nam = cells["C" + startRow].StringValue;
                        bc.ID_HinhThucKhenThuong = (short)GetDanhMuc_ID("HinhThucKhenThuong", cells["D" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu KhenThuong: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.SoQuyetDinh = cells["E" + startRow].StringValue;
                        bc.NgayRaQuyetDinh = cells["F" + startRow].StringValue;
                        bc.ID_NoiBanHanh = (short)GetDanhMuc_ID("NoiBanHanh", cells["G" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu KhenThuong: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }

                        if (bc.CanBo_KhenThuong_ID == 0)
                        {
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Thêm mới KhenThuong số: " + bc.SoQuyetDinh) });

                            db.CanBo_KhenThuongs.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu KhenThuong: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }

                #endregion
                #region Import KyLuat
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["KyLuat"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu KyLuat") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["C" + startRow].StringValue + "" == "")
                        {
                            startRow++;
                            continue;
                        }


                        var bc = new CanBo_KyLuat();
                        if (cells["A" + startRow].StringValue != "")
                        {
                            var id = cells["A" + startRow].IntValue;
                            bc = db.CanBo_KyLuats.FirstOrDefault(p => p.CanBo_KyLuat_ID == id);
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Sửa dữ liệu KyLuat: " + id) });
                        }
                        else
                        {
                            bc.CanBo_KyLuat_ID = 0;
                            bc.ID_CanBo = int.Parse(canbo.CanBo_ID);

                        }

                        bc.SoQuyetDinh = cells["C" + startRow].StringValue;
                        bc.NgayRaQuyetDinh = cells["D" + startRow].StringValue;
                        bc.ID_HinhThucKyLuat = (short)GetDanhMuc_ID("HinhThucKyLuat", cells["E" + startRow].StringValue, out err);
                        if (err != "")
                        {
                            importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu KyLuat: \"{0}\"", err) });
                            startRow++;
                            if (db.Transaction != null) db.Transaction.Rollback();
                            continue;
                        }
                        bc.NguyenNhan = cells["F" + startRow].StringValue;
                        bc.CoQuanRaQuyetDinh = cells["G" + startRow].StringValue;
                        bc.TuNgay = cells["H" + startRow].StringValue;
                        bc.DenNgay = cells["I" + startRow].StringValue;

                        if (bc.CanBo_KyLuat_ID == 0)
                        {
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Thêm mới Kỷ luật số: " + bc.SoQuyetDinh) });

                            db.CanBo_KyLuats.InsertOnSubmit(bc);
                            db.SubmitChanges();
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu KyLuat: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }

                #endregion
                #region Import QTluongCS
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["QTluongCS"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu QTluongCS") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["A" + startRow].Value + "" == "" || cells["J" + startRow].StringValue + "" == "" || cells["J" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var id = cells["A" + startRow].IntValue;
                        var bc = db.CanBo_BienDongLuongCungs.FirstOrDefault(p => p.CanBo_BienDongLuongCung_ID == id);

                        if (bc != null)
                        {
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Sửa dữ liệu QTluongCS: " + id) });
                            var heSoImport = Math.Round((decimal)cells["G" + startRow].FloatValue, 2);

                            if ((float)heSoImport != bc.HeSoLuong)
                            {
                                bc.SoQuyetDinh = cells["D" + startRow].StringValue;
                                if (cells["C" + startRow].Value + "" != "")
                                {
                                    DateTime ngayQd;
                                    if (DateTime.TryParseExact(cells["C" + startRow].Value + "", "d/M/yyyy", null,
                                        DateTimeStyles.None, out ngayQd))
                                    {
                                        bc.NgayQuyetDinh = ngayQd;
                                    }
                                }
                                bc.HeSoLuong = (float)heSoImport;

                                DateTime ngayHuong;
                                if (DateTime.TryParseExact(cells["H" + startRow].Value + "", "d/M/yyyy", null,
                                    DateTimeStyles.None, out ngayHuong))
                                {
                                    bc.HuongTuNgay = ngayHuong;
                                }

                                DateTime ngayGiuBac;
                                if (DateTime.TryParseExact(cells["I" + startRow].Value + "", "d/M/yyyy", null,
                                    DateTimeStyles.None, out ngayGiuBac))
                                {
                                    bc.NgayGiuBac = ngayGiuBac;
                                }

                                if (cells["J" + startRow].StringValue != bc.CanBo_BienDongLuongCung_ID + "" && cells["J" + startRow].StringValue + "" != "")
                                {
                                    bc.TenVanban = cells["J" + startRow].StringValue +
                                                   (cells["J" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");

                                    // process file Upload
                                    if (bc.TenVanban != "")
                                    {
                                        var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanban));

                                        if (fUpload.Exists)
                                        {
                                            fUpload.CopyTo(Path.Combine(UploadPath, "QuaTrinhLuongCB", bc.CanBo_BienDongLuongCung_ID + ""),
                                                true);
                                        }
                                    }
                                    // process file Upload
                                }
                            }
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu QTluongCS: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }

                #endregion
                #region Import QTluongCD
                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                sheet = workbook.Worksheets["QTluongCD"];
                cells = sheet.Cells;

                importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Lưu dữ liệu QTluongCD") });

                try
                {
                    startRow = 6;
                    while (startRow <= cells.MaxDataRow + 1)
                    {
                        if (cells["A" + startRow].Value + "" == "" || cells["K" + startRow].StringValue + "" == "" || cells["K" + startRow].StringValue == cells["A" + startRow].StringValue)
                        {
                            startRow++;
                            continue;
                        }

                        var id = cells["A" + startRow].IntValue;
                        var bc = db.CanBo_BienDongLuongChucDanhs.FirstOrDefault(p => p.CanBo_BienDongLuongChucDanh_ID == id);

                        if (bc != null)
                        {
                            //importWorker.ReportProgress(0, new string[] { item.Text, "status-away.png", string.Format("Sửa dữ liệu QTluongCS: " + id) });
                            var heSoImport = Math.Round((decimal)cells["I" + startRow].FloatValue, 2);

                            if ((float)heSoImport != bc.TongHeSo)
                            {
                                bc.SoQuyetDinh = cells["D" + startRow].StringValue;
                                if (cells["C" + startRow].Value + "" != "")
                                {
                                    DateTime ngayQd;
                                    if (DateTime.TryParseExact(cells["C" + startRow].Value + "", "d/M/yyyy", null,
                                        DateTimeStyles.None, out ngayQd))
                                    {
                                        bc.NgayQuyetDinh = ngayQd;
                                    }
                                }

                                bc.TongHeSo = (double)heSoImport;

                                DateTime ngayHuong;
                                if (DateTime.TryParseExact(cells["I" + startRow].Value + "", "d/M/yyyy", null,
                                    DateTimeStyles.None, out ngayHuong))
                                {
                                    bc.NgayApDung = ngayHuong;
                                }

                                if (cells["H" + startRow].Value + "" != "")
                                {
                                    bc.PhanTramPhuCapChucVu = cells["H" + startRow].FloatValue;
                                }

                                if (cells["K" + startRow].StringValue != bc.CanBo_BienDongLuongChucDanh_ID + "" && cells["K" + startRow].StringValue + "" != "")
                                {
                                    bc.TenVanBan = cells["K" + startRow].StringValue +
                                                   (cells["K" + startRow].StringValue.EndsWith(".pdf") ? "" : ".pdf");

                                    // process file Upload
                                    if (bc.TenVanBan != "")
                                    {
                                        var fUpload = new FileInfo(Path.Combine(txtFolder.Text, item.Text, bc.TenVanBan));

                                        if (fUpload.Exists)
                                        {
                                            fUpload.CopyTo(Path.Combine(UploadPath, "QuaTrinhLuongCD", bc.CanBo_BienDongLuongChucDanh_ID + ""),
                                                true);
                                        }
                                    }
                                    // process file Upload
                                }
                            }
                        }

                        startRow++;
                    }

                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Lỗi khi lưu dữ liệu QTluongCD: \"{0}\"", ex.Message) });
                    err = "";
                    if (db.Transaction != null) db.Transaction.Rollback();
                    continue;
                }

                #endregion

                if (importWorker.CancellationPending)
                {
                    importWorker.ReportProgress(0, new string[] { item.Text, "status-busy.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                    break;
                }

                importWorker.ReportProgress(0, new string[] { item.Text, "status.png", string.Format("Kết thúc Import thư mục: \"{0}\"", item.Text) });
                importWorker.ReportProgress(0, new string[] { item.Text, "status.png", string.Format("--------------------------------", item.Text) });
            }
        }

        #endregion
    }
}
