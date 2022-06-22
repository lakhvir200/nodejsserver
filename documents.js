using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
//using System.Data.Sql;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using User;
using System.Threading;
using ClosedXML.Excel;


//using Microsoft.Office.Interop.Excel;


namespace Equipment_project_Accessdata
{

    public partial class AddEquip : Form
    {
        public static String connString;
        bool Equipment = true, Repair = false, Purchase = false, Consumption = false,
            Maintenance = false, Pending = false, DeptName = false, EquipName = false, Stock = false, ReminderTasks = false;
        string sql;
        private int rowIndex = 0;
        int id = 0;





        //BackgroundWorker bw = new BackgroundWorker
        //{
        //    WorkerReportsProgress = true,
        //    WorkerSupportsCancellation = true
        //};

        public AddEquip()
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer1.Interval = 1000;

            lbltimecheck.Text = "5";
            timerReminder.Start();
        }

        //private DataTable dtAddEquip = new DataTable();

        private void AddEquip_Load(object sender, EventArgs e)
        {
            ButtonClick();
            ReminderLoad();
            //Screen resolution setting
            Rectangle resolutionRect = System.Windows.Forms.Screen.FromControl(this).Bounds;
            if (this.Width >= resolutionRect.Width || this.Height >= resolutionRect.Height)
            {
                this.WindowState = FormWindowState.Maximized;

            }

        }


        private void BtnAttchBill_Click(object sender, EventArgs e)
        {
            ButtonClick();
            //string Equip_ID = txtid.Text;
            //Directory.CreateDirectory(@Doc_Path + "\\" + Equip_ID + "\\");

            // lblstatus.Visible = false;
            // groupBoxEquip.Visible = true;
            // groupBoxPurc.Visible = false;
            // groupBoxConsumption.Visible = false;
            // groupBoxRepair.Visible = false;
            // groupBoxMaint.Visible = false;
            // groupBoxStock.Visible = false;

            // EquipGrid.Location = new Point(170, 100);
            // EquipGrid.Size = new Size(1100, 520);
            // groupBoxEquip.Location = new Point(230, 40);
            //groupBox3.Location = new Point(149, 625);

            //load();           

        }

        private void EquipGrid_DoubleClick(object sender, EventArgs e)
        {
            //int i = this.EquipGrid.CurrentRow.Cells[0].RowIndex;


            UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();

            UpdateEquipNew UE = new UpdateEquipNew();
            UE.Show();
            this.Visible = true;
            // LoadDataIntoEquipGridView();

        }
        private void load()
        {
            SqlCommand cmdDept, cmdEquip, cmdHosp, cmdID;
            SqlDataReader dr;
            String hospital = cmbHosp.Text;

            using (SqlConnection con = new SqlConnection(connString))
            {
                cmdDept = new SqlCommand("Select distinct DEPARTMENT from dept_details order  by DEPARTMENT  asc", con);
                cmdEquip = new SqlCommand("Select distinct Equip_Name from Equip_Names order by Equip_Name asc", con);
                cmdHosp = new SqlCommand("Select distinct Equip_Name from Equipment order by  Equip_Name asc", con);
                cmdID = new SqlCommand("Select distinct Equipment_ID from new_EQUIPMENT_INFO order by  Equipment_ID asc ", con);
                //  using (SqlCommand cmd = new SqlCommand(cmdString, con))
                {
                    //----adding equipment names to combocon.Open();
                    con.Open();
                    dr = cmdEquip.ExecuteReader();
                    cmbEquip.Items.Add("All");
                    while (dr.Read()) //loop
                        cmbEquip.Items.Add(dr["Equip_Name"].ToString());
                    con.Close();

                    //......adding department names to combodepartment
                    con.Open();
                    dr = cmdDept.ExecuteReader();
                    cmbDept.Items.Add("All");
                    while (dr.Read()) //loop
                        cmbDept.Items.Add(dr["DEPARTMENT"].ToString());
                    con.Close();
                    //adding id in combo id
                    con.Open();
                    dr = cmdID.ExecuteReader();
                    while (dr.Read()) //loop
                        cmbID.Items.Add(dr["Equipment_ID"].ToString());
                    con.Close();

                    string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,Category,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");
                    sql = p + " where unit_Name ='Beas Hosp' and equip_status ='OK' order by equipment_id ";
                    con.Open();

                    SqlDataAdapter dataadapter = new SqlDataAdapter(sql, con);
                    DataSet ds = new DataSet();
                    DataTable tbl = new DataTable();
                    dataadapter.Fill(tbl);
                    EquipGrid.DataSource = tbl;

                    con.Close();
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    //Group box
                    //groupBoxEquip.Visible = true;
                    //groupBoxPurc.Visible = false;
                    //groupBoxConsumption.Visible = false;
                    //groupBoxRepair.Visible = false;
                    //groupBoxMaint.Visible = false;
                    Equipload();

                }
            }
        }
        private void showData()
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                String dept = cmbDept.Text;
                String status = cmbStatus.Text;
                String equip = cmbEquip.Text;
                String hospital = cmbHosp.Text;
                String ID = cmbID.Text;
                //EquipGrid.Columns["view"].DefaultCellStyle.BackColor = Color.Gold;


                try
                {
                    string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,Category,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");

                    // string p = ("Select ID,EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");

                    if (dept.Equals("All") && equip.Equals("All") && hospital.Equals("All") && status.Equals("All"))//0000
                        sql = p + "ORDER BY EQUIPMENT_ID";


                    else if (dept.Equals("All") && equip.Equals("All") && hospital.Equals("All") && !status.Equals("All"))//0001
                        sql = p + "where equip_status = '" + status + "'  order by equipment_id";

                    else if (dept.Equals("All") && equip.Equals("All") && !hospital.Equals("All") && status.Equals("All"))//0010
                        sql = p + "where unit_name ='" + hospital + "' order by equipment_id";

                    else if (dept.Equals("All") && equip.Equals("All") && !hospital.Equals("All") && !status.Equals("All"))//0011
                        sql = p + "where unit_name ='" + hospital + "' and equip_status ='" + status + "' order by equipment_id";

                    else if (dept.Equals("All") && !equip.Equals("All") && hospital.Equals("All") && status.Equals("All"))//0100
                        sql = p + "  where Equipment_Name ='" + equip + "'  order by equipment_id";

                    else if (dept.Equals("All") && !equip.Equals("All") && hospital.Equals("All") && !status.Equals("All"))//0101
                        sql = p + "where Equipment_Name ='" + equip + "' and equip_status ='" + status + "' order by equipment_id";

                    else if (dept.Equals("All") && !equip.Equals("All") && !hospital.Equals("All") && status.Equals("All"))//0110
                        sql = p + " where Equipment_Name ='" + equip + "' and  unit_name ='" + hospital + "' order by equipment_id  ";

                    else if (dept.Equals("All") && !equip.Equals("All") && !hospital.Equals("All") && !status.Equals("All"))//0111
                        sql = p + "  where Equipment_Name ='" + equip + "'  and unit_name ='" + hospital + "' and equip_status ='" + status + "'order by equipment_id";

                    else if (!dept.Equals("All") && equip.Equals("All") && hospital.Equals("All") && status.Equals("All"))//1000
                        sql = p + " where department ='" + dept + "' order by equipment_id  ";

                    else if (!dept.Equals("All") && equip.Equals("All") && hospital.Equals("All") && !status.Equals("All"))//1001
                        sql = p + " where department ='" + dept + "'and equip_status ='" + status + "' order by equipment_id  ";

                    else if (!dept.Equals("All") && equip.Equals("All") && !hospital.Equals("All") && status.Equals("All"))//1010
                        sql = p + " where department ='" + dept + "'  and unit_name ='" + hospital + "' order by equipment_id  ";

                    else if (!dept.Equals("All") && equip.Equals("All") && !hospital.Equals("All") && !status.Equals("All"))//1011
                        sql = p + " where  department ='" + dept + "' and  unit_name ='" + hospital + "' and equip_status ='" + status + "'  order by equipment_id  ";

                    else if (!dept.Equals("All") && !equip.Equals("All") && hospital.Equals("All") && status.Equals("All"))//1100
                        sql = p + " where  department ='" + dept + "' and Equipment_Name ='" + equip + "' order by equipment_id  ";

                    else if (!dept.Equals("All") && !equip.Equals("All") && hospital.Equals("All") && !status.Equals("All"))//1101
                        sql = p + " where  department ='" + dept + "' and Equipment_Name ='" + equip + "' and equip_status ='" + status + "'  order by equipment_id  ";

                    else if (!dept.Equals("All") && !equip.Equals("All") && !hospital.Equals("All") && status.Equals("All"))//1110
                        sql = p + " where  department ='" + dept + "' and Equipment_Name ='" + equip + "' and  unit_name ='" + hospital + "' order by equipment_id  ";

                    else if (!dept.Equals("All") && !equip.Equals("All") && !hospital.Equals("All") && !status.Equals("All"))//1111
                        sql = p + " where  department ='" + dept + "' and Equipment_Name ='" + equip + "' and  unit_name ='" + hospital + "' and equip_status ='" + status + "'  order by equipment_id  ";

                    con.Open();

                    SqlDataAdapter dataadapter = new SqlDataAdapter(sql, con);
                    //DataSet ds = new DataSet();
                    DataTable tbl = new DataTable();
                    dataadapter.Fill(tbl);
                    EquipGrid.DataSource = tbl;
                    //dataadapter.Fill(ds, "new_EQUIPMENT_INFO");
                    //EquipGrid.DataSource = ds;
                    //EquipGrid.DataMember = "new_EQUIPMENT_INFO";
                    EquipGrid.Refresh();
                    con.Close();
                    Equipload();
                }
                catch
                {

                }
                //EquipGrid.Columns[0].Width = 50;
                //EquipGrid.Columns[1].Width = 90;
                //EquipGrid.Columns[2].Width = 120;
                //EquipGrid.Columns[3].Width = 400;
                //EquipGrid.Columns[4].Width = 100;
                //EquipGrid.Columns[5].Width = 100;
            }
        }

        private void cmbStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            showData();
            string qty = Convert.ToString(EquipGrid.Rows.Count);
            lblCount.Text = qty;

        }

        private void cmbDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            showData();
            string qty = Convert.ToString(EquipGrid.Rows.Count);
            lblCount.Text = qty;
            Equipload();


        }

        private void cmbEquip_SelectedIndexChanged(object sender, EventArgs e)
        {
            showData();
            string qty = Convert.ToString(EquipGrid.Rows.Count);
            lblCount.Text = qty;
        }

        private void cmbHosp_SelectedIndexChanged(object sender, EventArgs e)
        {

            //if (txtEquipSearch.Text=="")

            showData();
            //else
            //{
            //   EquipSearch();
            //}

            string qty = Convert.ToString(EquipGrid.Rows.Count);
            lblCount.Text = qty;
        }

        private void cmbID_SelectedIndexChanged(object sender, EventArgs e)
        {

            using (SqlConnection con = new SqlConnection(connString))
            {
                //    String dept = cmbDept.Text;
                //    String status = cmbStatus.Text;
                //    String equip = cmbEquip.Text;
                //    String hospital = cmbHosp.Text;
                String ID = cmbID.Text;
                string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");
                sql = p + "where Equipment_ID = '" + ID + "'";
                con.Open();

                SqlDataAdapter dataadapter = new SqlDataAdapter(sql, con);
                DataTable tbl = new DataTable();
                dataadapter.Fill(tbl);
                EquipGrid.DataSource = tbl;
                //DataSet ds = new DataSet();
                //dataadapter.Fill(ds, "new_EQUIPMENT_INFO");
                //EquipGrid.DataSource = ds;
                //EquipGrid.DataMember = "new_EQUIPMENT_INFO";
                EquipGrid.Refresh();
                con.Close();
                Equipload();
                //string qty = Convert.ToString(EquipGrid.Rows.Count);

                //lblCount.Text = qty;


            }
        }




        private void btnExport_Click(object sender, EventArgs e)
        {
            ExporttoXLS();
            progressBar.Visible = true;
            lblstatus.Visible = true;
            groupBox3.Visible = false;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();


        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 1; i <= 100; i++)
            {
                // ExporttoXLS();
                // Wait 50 milliseconds.  
                Thread.Sleep(10);
                // Report progress.  
                backgroundWorker1.ReportProgress(i);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar   
            progressBar.Value = e.ProgressPercentage;
            // Set the text.  
            this.Text = e.ProgressPercentage.ToString();
            lblstatus.Text = string.Format("Processing...{0}%", e.ProgressPercentage);
            progressBar.Update();


        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                lblstatus.Text = "Data is successfully exported";
                MessageBox.Show("Data is exported to EXLS successfully");
                progressBar.Visible = false;
                lblstatus.Visible = false;
                if (Equipment)
                {
                    groupBox3.Visible = true;
                }
                else
                    groupBox3.Visible = false;
            }

        }



        private void ExporttoXLS()
        {
            if (backgroundWorker1.IsBusy)
                return;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.FileName = "";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlxs";

            if (sfd.ShowDialog() != DialogResult.Cancel)
            {
                // Using Excel = Microsoft.Office.Interop.Excel;
                // creating Excel Application
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                //  creating new WorkBook within Excel application
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                for (int i = 1; i < EquipGrid.Columns.Count + 1; i++)
                {
                    app.Cells[1, i] = EquipGrid.Columns[i - 1].HeaderText;
                }
                //progressBar.Minimum = 0;
                //progressBar.Maximum = EquipGrid.Rows.Count;
                //storing Each row and column value to excel sheet
                for (int i = 0; i < EquipGrid.Rows.Count; i++)
                {
                    for (int j = 0; j < EquipGrid.Columns.Count; j++)
                    {
                        app.Cells[i + 2, j + 1] = EquipGrid.Rows[i].Cells[j].Value.ToString();

                    }
                    //progressBar.Value = i;
                    //decimal percent = (progressBar.Value / progressBar.Maximum);
                    //percent = percent * 100;
                    //lblstatus.Text = Convert.ToString(percent);
                }

                app.ActiveWorkbook.SaveCopyAs(sfd.FileName.ToString());
                app.ActiveWorkbook.Saved = true;
                app.ActiveWorkbook.Saved = true;
                MessageBox.Show("Data is exported to EXLS successfully");

                app.Quit();

            }
        }


        private void btnRefr_Click(object sender, EventArgs e)
        {
            showData();
            string qty = Convert.ToString(EquipGrid.Rows.Count);
            lblCount.Text = qty;
        }

        private void viewform()
        {

            Equip_Detail_new.connString = connString;

            int i = EquipGrid.SelectedCells[0].RowIndex;

            Equip_Detail_new.equipid = EquipGrid.Rows[i].Cells[0].Value.ToString();
            Equip_Detail_new.equipname = EquipGrid.Rows[i].Cells[1].Value.ToString();
            Equip_Detail_new.speci = EquipGrid.Rows[i].Cells[2].Value.ToString();
            Equip_Detail_new.dept = EquipGrid.Rows[i].Cells[3].Value.ToString();
            Equip_Detail_new.purchdate = EquipGrid.Rows[i].Cells[4].Value.ToString();
            Equip_Detail_new.price = EquipGrid.Rows[i].Cells[5].Value.ToString();
            Equip_Detail_new.status = EquipGrid.Rows[i].Cells[6].Value.ToString();

            Equip_Detail_new VE = new Equip_Detail_new();
            VE.Show();
            this.Visible = true;

        }



        private void EquipGrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    ContextMenu m = new ContextMenu();
            //    //m.MenuItems.Add(new MenuItem("Update"));
            //    m.MenuItems.Add(new MenuItem("View"));
            //    // m.MenuItems.Add(new MenuItem("Paste"));

            //    int currentMouseOverRow = EquipGrid.HitTest(e.X, e.Y).RowIndex;

            //    if (currentMouseOverRow >= 0)
            //    {
            //        // m.MenuItems.Add(new MenuItem(string.Format("Do something to row {0}", currentMouseOverRow.ToString())));
            //        m.MenuItems.Add(new MenuItem(string.Format("Do something to row {0}", currentMouseOverRow.ToString())));
            //    }
            //    viewform();
            //}
        }



        private void btnReset_Click(object sender, EventArgs e)
        {

            //foreach (Control c in groupBox1.Controls)
            //{
            //    if (c is ComboBox)
            //    {
            //        c.Text = "";
            //    }
            //}
            load();

            //  groupBox2.Visible = false;

        }





        private void viewToolStripMenuItem_Click(object sender, EventArgs e)
        {

            New_Equipment_Addition VE = new New_Equipment_Addition();
            VE.Show();
            this.Visible = true;

        }

        private void viewToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AddRepair BE = new AddRepair();
            BE.Show();
            this.Visible = true;
        }




        //private void EquipGrid_MouseDown(object sender, MouseEventArgs e)
        //{
        //    EquipGrid.MouseDown += new MouseEventHandler(this.EquipGrid_MouseClick);

        //}



        private void editToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // con = new SqlConnection();
            //groupBox3.Visible = false;
            //groupBox2.Visible = false;
            // groupBoxEquip.Visible = false;
            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Repair_Master   order by RDATE desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Repair_master");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Repair_master";
            }
        }

        private void EquipGrid_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Equipment)
            {
                try
                {
                    UpdateEquipNew.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    UpdateEquipNew UE = new UpdateEquipNew();
                    UE.Show();
                    this.Visible = true;
                }
                catch
                {

                }

            }
            else if (Repair)
            {
                int i = EquipGrid.SelectedCells[0].RowIndex;

                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();


                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;

            }
            else if (Pending)
            {
                int i = EquipGrid.SelectedCells[0].RowIndex;

                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();


                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;

            }
            else if (EquipName)
            {
                AdditionEquipList.connString = connString;
                AdditionEquipList UE = new AdditionEquipList();
                UE.Show();
                this.Visible = true;

            }

        }

        private void viewToolStripMenuItem2_Click(object sender, EventArgs e)
        {

            ButtonClick();
            Purchase = true;
            btnPurch.ForeColor = System.Drawing.Color.Red;
            groupBoxPurc.Visible = true;


            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Purchase order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Purchase");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Purchase";
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();
            }
        }

        private void viewToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ButtonClick();

            btnCons.ForeColor = System.Drawing.Color.Red;

            Consumption = true;
            groupBoxConsumption.Visible = true;


            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from consumption  order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "consumption");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "consumption";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Consumpload();
            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtEquipSearch_TextChanged(object sender, EventArgs e)
        {
            //using (SqlConnection con = new SqlConnection(connString))
            //{
            //    con.Open();
            //    SqlDataAdapter dataadapter = new SqlDataAdapter("select * from new_Equipment_info where Complete_specification like '" + txtEquipSearch.Text + "%'", con);
            //    DataTable dt = new DataTable();
            //    dataadapter.Fill(dt);
            //    EquipGrid.DataSource = dt;
            //    con.Close();
            //}
        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                try
                {
                    UpdateEquipNew.connString = connString;

                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    UpdateEquipNew UE = new UpdateEquipNew();
                    UE.Show();
                    this.Visible = true;
                    //String Desc = "";
                    //string Q = "select * from reminders where EquipId = '" + this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString()+ "' and ISReminderSet ";
                    //using (SqlConnection con = new SqlConnection(connString))
                    //{
                    //    con.Open();
                    //    SqlCommand CM = new SqlCommand(Q);
                    //    CM.Connection = con;
                    //    SqlDataReader Rd = CM.ExecuteReader();
                    //    if(Rd.HasRows)
                    //    {                          
                    //        while(Rd.Read())
                    //        {
                    //           // Desc = Rd[2].ToString();
                    //            Desc +="\n Equipment has A reminder set "+ Rd.GetString(2)+" Due date is " +Rd.GetDateTime(3).ToString();
                    //        }
                    //        MessageBox.Show(Desc);
                    //    }

                    //    con.Close();
                    //}


                }

                catch
                {

                }

            }
            else if (Repair)
            {
                AddRepair.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;

            }
            else if (Maintenance)
            {
                AddRepair.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;

            }
            else if (Pending)
            {
                AddRepair.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;

            }
            else if (EquipName)
            {
                AdditionEquipList.connString = connString;
                AdditionEquipList UE = new AdditionEquipList();
                UE.Show();
                this.Visible = true;

            }

            else if (DeptName)
            {
                AddNewDept.connString = connString;
                AddNewDept UE = new AddNewDept();
                UE.Show();
                this.Visible = true;

            }
            else if (Consumption)
            {
                EditConsumption.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                EditConsumption.ItemID = EquipGrid.Rows[i].Cells["eqid"].Value.ToString();
                EditConsumption.ID = EquipGrid.Rows[i].Cells["ID"].Value.ToString();
                EditConsumption UE = new EditConsumption();
                UE.Show();
                this.Visible = true;

            }

            else if (ReminderTasks)
            {
                ReminderUpdate.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                ReminderUpdate.Date = EquipGrid.Rows[i].Cells["DueDate"].Value.ToString();
                ReminderUpdate.ID = EquipGrid.Rows[i].Cells["ID"].Value.ToString();
                ReminderUpdate UE = new ReminderUpdate();
                UE.Show();
                this.Visible = true;

            }


        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddPurchase.connString = connString;
            AddPurchase UE = new AddPurchase();
            UE.Show();
            this.Visible = true;

        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int i = EquipGrid.SelectedCells[0].RowIndex;
            Equip_Detail_new.equipid = EquipGrid.Rows[i].Cells[0].Value.ToString();
            Equip_Detail_new.equipname = EquipGrid.Rows[i].Cells[1].Value.ToString();
            Equip_Detail_new.speci = EquipGrid.Rows[i].Cells[2].Value.ToString();
            Equip_Detail_new.dept = EquipGrid.Rows[i].Cells[3].Value.ToString();
            Equip_Detail_new.purchdate = EquipGrid.Rows[i].Cells[4].Value.ToString();
            Equip_Detail_new.price = EquipGrid.Rows[i].Cells[5].Value.ToString();
            Equip_Detail_new.status = EquipGrid.Rows[i].Cells[6].Value.ToString();


            Equip_Detail_new VE = new Equip_Detail_new();
            VE.Show();
            this.Visible = true;
        }

        private void editToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                int i = EquipGrid.SelectedCells[0].RowIndex;
                UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();

                UpdateEquipNew UE = new UpdateEquipNew();
                UE.Show();
                this.Visible = true;
                // LoadDataIntoEquipGridView();
            }

            catch
            {

            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("you want to Delete This Record", "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {

                // No row selected no delete....
                if (EquipGrid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("No row selected !");// show a message here to inform
                }

                // Prepare the command text with the parameter placeholder

                string sql = "DELETE FROM new_Equipment_info  WHERE ID=@rowID";
                using (SqlConnection con = new SqlConnection(connString))

                // Create the connection and the command inside a using block
                //using (SqlConnection myConnection = new SqlConnection("Data Source=Epic-LaptopWR;Initial Catalog=words;Integrated Security=True"))
                using (SqlCommand deleteRecord = new SqlCommand(sql, con))

                {
                    con.Open();

                    // if (EquipGrid.CurrentCell.RowIndex >0)
                    {
                        int rowIndex = EquipGrid.CurrentCell.RowIndex;
                        int selectedIndex = EquipGrid.SelectedRows[0].Index;
                        // gets the RowID from the first column in the grid
                        string rowID = Convert.ToString(EquipGrid[0, selectedIndex].Value);

                        deleteRecord.Parameters.AddWithValue("@rowID", rowID);//.Value.ToString();//= rowID;
                        deleteRecord.ExecuteNonQuery();
                        MessageBox.Show("Record Deleted Successfully");

                        // Remove the row from the grid
                        EquipGrid.Rows.RemoveAt(selectedIndex);
                        //EquipGrid.Rows.RemoveAt(EquipGrid.CurrentRow.Index);
                    }

                }// end of delete Button 
            }
            else
            {
                MessageBox.Show("Record not Deleted", "Delete Record", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
        }

        private void addToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ADDConsumption.connString = connString;
            ADDConsumption UE = new ADDConsumption();
            UE.Show();
            this.Visible = true;

        }

        private void deleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to Delete Duplicate Record", "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string sql = ("WITH CTE AS (SELECT *, RN = ROW_NUMBER()OVER(PARTITION BY[DOC_NO],[DRUG_SRNO]  ORDER BY DOC_NO desc)FROM[Equipment].[dbo].[Purchase]) Delete CTE where RN > 1");

                using (SqlConnection con = new SqlConnection(connString))

                using (SqlCommand deleteRecord = new SqlCommand(sql, con))

                {
                    con.Open();

                    deleteRecord.ExecuteNonQuery();
                    MessageBox.Show("Record Deleted Successfully!");

                }
            }
            else
            {
                MessageBox.Show("Record not Deleted", "Delete Record", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            ButtonClick();
            Purchase = true;
            groupBoxPurc.Location = new Point(300, 10);
            //Equipment = Repair = Consumption = Maintenance = Pending = Stock== false;            
            //groupBox3.Visible = false;
            //groupBoxConsumption.Visible = false;
            //groupBoxEquip.Visible = false;
            //groupBoxPurc.Visible = true;
            //groupBoxRepair.Visible = false;
            //groupBoxMaint.Visible = false;

            // groupBoxEquip.Visible = false;// con = new SqlConnection();
            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Purchase order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Purchase");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Purchase";
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();
            }


        }
        private void addEquipmentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //groupBoxEquip.Visible = true;
            //groupBoxPurc.Visible = false;
            //groupBoxConsumption.Visible = false;
            //groupBoxRepair.Visible = false;
            //load();
        }
        private void addRepairToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }
        private void purchaseToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void consuptionToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btnSearchPurc_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                // SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DRUG_NM  LIKE '%" + txtSearchPurc.Text + "%'", con);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where CONCAT ([DRUG_NM],[SUPP_NAME],DATEPART(yyyy,DOC_DT))  LIKE '%" + txtSearchPurc.Text + "%'", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();

                /*CONCAT([Complete_Specification], [Equipment_Name],[EQUIPMENT_ID], DATEPART(yyyy,DATE_OF_PURCHASE)) LIKE '%" + txtSearch.Text + "%' ";


                BindingSource bs = new BindingSource();
                bs.DataSource = EquipGrid.DataSource;
               // bs.Filter = "DRUG_NM like '%" + txtSearchPurc.Text + "%'";
                EquipGrid.DataSource = bs;
                (bs.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '%{0}%'", txtSearchPurc.Text);
                */
            }
        }

        private void btnSearchCons_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                //select * from bio_db.daily_data2 where Date between '" + dtFrom.ToString("MM/dd/yyyy")+ "' and '" + dtTo.ToString("MM/dd/yyyy") + "' ", mcon)
                //where Date between '" + datefrom.Value.ToString() + "' and '" + dateto.Value.ToString() + "' ", mcon);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM consumption Where CONCAT ([DRUG_NM],[SUPP_NAME],DATEPART(yyyy,DOC_DT))  LIKE '%" + txtSearchCons.Text + "%'", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Consumpload();

            }
        }

        private void cmbIdRepair_SelectedIndexChanged(object sender, EventArgs e)
        {

            String ID = cmbIdRepair.Text;

            using (SqlConnection con = new SqlConnection(connString))
            {
                string p = ("Select Repair_Master.RID,Repair_Master.Rdate,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where Eq_ID='" + ID + "'order by Rdate desc ");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Repair_master");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Repair_master";
            }
        }

        private void cmbIdPurc_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbIdCons_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {

                String ID = cmbID.Text;
                string p = ("Select* From consumption where eqid='" + cmbIdCons.Text + "'");

                con.Open();

                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "consumption");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "consumption";
                EquipGrid.Refresh();
                con.Close();
                Consumpload();

            }
        }

        private void btnRepair_Click(object sender, EventArgs e)
        {
            ButtonClick();
            //btnRepair.Text = "Repair >>";
            //btnEquip.TextAlign = ContentAlignment.MiddleRight;
            btnRepair.ForeColor = System.Drawing.Color.Red;
            groupBoxRepair.Visible = true;

            //EquipGrid.DataSource = null;
            //EquipGrid.Refresh();
            //EquipGrid.Columns.Clear();
            //EquipGrid.Rows.Clear();


            //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = Reminder=false;
            Repair = true;



            using (SqlConnection con = new SqlConnection(connString))
            {
                //SELECT Orders.OrderID, Customers.CustomerName FROM Orders iNNER JOIN Customers ON Orders.CustomerID = Customers.CustomerID;
                string p = ("Select TOP 500 Repair_Master.RID,Repair_Master.Rdate,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id order by Rdate desc ");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataTable Dt = new DataTable();
                dataadapter.Fill(Dt);
                //EquipGrid.Refresh();
                //EquipGrid.Rows.Clear();              

                EquipGrid.DataSource = Dt;
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Repairload();



            }
        }

        private void btnEquip_Click(object sender, EventArgs e)
        {
            ButtonClick();
            Equipment = true;
            groupBoxEquip.Visible = true;
            //groupBox3.Visible = true;
            //groupBox3.Location = new Point(149, 625);
            btnEquip.ForeColor = System.Drawing.Color.Red;
            load();
            //btnEquip.Text = " Equipment>>";
            //btnEquip.TextAlign = ContentAlignment.MiddleRight;
            //btnEquip.ForeColor = System.Drawing.Color.Red;


            //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock =Reminder= false;


            //groupBoxMaint.Visible = false;
            //groupBoxStock.Visible = false;
            //groupBoxEquip.Location = new Point(230, 40);

            //EquipGrid.Columns[6].Width = 80;
            //EquipGrid.Columns[7].Width = 250;
            //EquipGrid.Columns[8].Width = 100;




        }

        private void btnView_Click(object sender, EventArgs e)
        {

            if (Equipment)
            {
                try

                {
                    Equip_Detail_new.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    Equip_Detail_new UE = new Equip_Detail_new();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }


            }
            else if (Repair)
            {
                AddRepair.connString = connString;
               
                AddRepair VE = new AddRepair();
                VE.Show();
                this.Visible = true;

            }
            else if (Purchase)
            {
            }
            else if (Consumption)
            {

            }
            else if (Maintenance)
            {
                Maint.connString = connString;
                Maint VE = new Maint();
                VE.Show();
                this.Visible = true;

            }
            else if (Pending)
            {


            }

            else if (ReminderTasks)
            {

                ReminderCheck();
            }


        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                try
                {
                    UpdateEquipNew.connString = connString;

                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    UpdateEquipNew UE = new UpdateEquipNew();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }

            }
            else if (Repair)
            {

            }
            else if (Purchase)
            {

            }
            else if (Consumption)
            {

            }
            else if (Pending)
            {
                AddRepair.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;

                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();


                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;
            }

        }

        private void btnEqSearch_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                // SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DRUG_NM  LIKE '%" + txtSearchPurc.Text + "%'", con);
                if (cmbHosp.Text.Equals("All"))
                {
                    string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");
                    sql = p + " where equip_status ='OK' and COMPLETE_SPECIFICATION LIKE '%" + txtEquipSearch.Text + "%' order by equipment_id ";
                }

                else
                {
                    string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");
                    sql = p + " where equip_status ='OK' and Unit_Name ='" + cmbHosp.Text + "' and COMPLETE_SPECIFICATION LIKE '%" + txtEquipSearch.Text + "%' order by equipment_id ";
                    // sql = p + " where equip_status ='OK' and Unit_Name ='Beas Hosp' order by equipment_id ";
                    // string p = ("Select EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");
                }

                SqlDataAdapter Sda = new SqlDataAdapter(sql, con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                EquipGrid.Columns[0].Width = 50;
                EquipGrid.Columns[1].Width = 90;
                EquipGrid.Columns[2].Width = 120;
                EquipGrid.Columns[3].Width = 400;
                EquipGrid.Columns[4].Width = 100;
                EquipGrid.Columns[5].Width = 100;
            }
        }

        private void btnPend_Click(object sender, EventArgs e)
        {
            ButtonClick();

            btnPend.ForeColor = System.Drawing.Color.Red;
            Pending = true;


            //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = false;


            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = "SELECT t1.[RID],t1.[RDATE],t1.[EQ_ID],new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification," +
                   "t1.[DEPT],t1.[REPAIR_MAINT],t1.[ACTION_TAKEN],t1.[SAPRES],t1.[ATT_BY],t1.[STATUS],t1.[RNO]FROM([Equipment].[dbo].[REPAIR_MASTER] t1 inner join[Equipment].[dbo].new_EQUIPMENT_INFO ON " +
                   "t1.EQ_ID = New_Equipment_info.Equipment_Id inner join (select Eq_ID, max(RDate) as MaxDate FROM [Equipment].[dbo].[REPAIR_MASTER] group by Eq_ID) tm on t1.EQ_ID=tm.EQ_ID and t1.RDATE =tm.MaxDate)" +
                   "where t1.STATUS like '%Pending%' ORDER BY t1.RDATE desc";


                SqlCommand cmd = new SqlCommand(p, con);
                /****** Script for SelectTopNRows command from SSMS  ******/

                // select filename,status,max_date = max(dates)from some_table tgroup by filename , status having status = '<your-desired-status-here>'

                // string p = ("select Status, Rdate from Repair_master  group by  having status='Pending'");

                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Repair_master");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Repair_master";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                // EquipGrid.Columns.("Phone").Width = 30
                EquipGrid.Columns[0].Width = 60;
                EquipGrid.Columns[1].Width = 90;
                EquipGrid.Columns[2].Width = 100;
                EquipGrid.Columns[3].Width = 150;
                EquipGrid.Columns[4].Width = 300;
                EquipGrid.Columns[5].Width = 100;
                EquipGrid.Columns[6].Width = 80;
                EquipGrid.Columns[7].Width = 350;
                EquipGrid.Columns[8].Width = 100;
                EquipGrid.Columns[9].Width = 50;
                //EquipGrid.Columns[10].Width = 200;

            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                // SELECT Repair FROM new_equipment_info GROUP BY Equipment_Name HAVING COUNT(*) > 1;

                // SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DRUG_NM  LIKE '%" + txtSearchPurc.Text + "%'", con);
                string p = ("Select Equipment_Name,  COUNT(*) AS Times_of_Repairs from New_Equipment_info INNER JOIN Repair_Master ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id Where Repair_Master.Rdate between '" + dateTimePickerRepairFrom.Value.ToString("yyyy/MM/dd") + "' and '" + dateTimePickerRepairTo.Value.ToString("yyyy/MM/dd") + "'GROUP BY Equipment_Name HAVING COUNT(*) > 1  ");// Repair_Master.Eq_ID,new_equipment_info.Equipment_Name from Repair_master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id GROUP BY Equipment_Name HAVING COUNT(*) > 1 ");
                SqlDataAdapter Sda = new SqlDataAdapter(p, con);// Where Rdate between '" + dateTimePickerRepairFrom.Value.ToString("yyyy/MM/dd") + "' and '" + dateTimePickerRepairTo.Value.ToString("yyyy/MM/dd") + "'", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;

                //COUNT(*)AS DUPLICATES_COUNT,
            }
        }

        private void btnMaint_Click(object sender, EventArgs e)
        {
            ButtonClick();
            btnMaint.ForeColor = System.Drawing.Color.Red;
            Maintenance = true;
            groupBoxMaint.Visible = true;
            string Cat = cmbCategory.Text;
            SqlCommand cmdCat;
            SqlDataReader dr;
            //Convert combobox Month Name to number
            //int month = DateTime.Parse("1." + CmbMonth.Text + "").Month;
            //int Year = DateTime.Parse("1." + CmbYear.Text + "").Year;
            // string Category = cmbCategory.Text;

            using (SqlConnection con = new SqlConnection(connString))
            {
                cmdCat = new SqlCommand("Select Category from Category order  by Category  asc", con);

                string p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name = 'Beas Hosp'and Repair_Master.Repair_maint = 'Maint.' order by Rdate desc ");
                
                con.Open();
                SqlDataAdapter Sda = new SqlDataAdapter(p, con);
                DataTable Dt = new DataTable();
                EquipGrid.DataSource = Dt;
                Sda.Fill(Dt);
                con.Close();
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;


                con.Open();
                dr = cmdCat.ExecuteReader();
                cmbCategory.Items.Add("All");
                while (dr.Read()) //loop
                    cmbCategory.Items.Add(dr["Category"].ToString());
                con.Close();
                Maintload();
                try
                {
                    //clear all data from combobox
                    CmbMonth.Items.Clear();
                    //add default item
                    CmbMonth.Items.Add("Select");
                    //fill array from System.Globalization.DateTimeFormatInfo.InvariantInfo
                    var Months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames;
                    //loop array for add items
                    foreach (string sm in Months)
                    {
                        if (sm != "")
                            CmbMonth.Items.Add(sm);
                    }
                    //set selected item for display on startup
                    CmbMonth.Text = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                //Add year in Combobox

                try
                {
                    //clear all data from combobox
                    CmbYear.Items.Clear();
                    //add default item
                    CmbYear.Items.Add("Select");
                    //loop array for add items
                    for (int i = DateTime.Now.Year; i < DateTime.Now.Year + 15; i++)
                    {
                        CmbYear.Items.Add(i);
                    }
                    //set selected item for display on startup
                    CmbYear.Text = DateTime.Now.Year.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                New_Equipment_Addition.connString = connString;
                New_Equipment_Addition VE = new New_Equipment_Addition();
                VE.Show();
                this.Visible = true;
            }
            else if (Repair)
            {
                AddRepair.connString = connString;
                AddRepair BEP = new AddRepair();
                BEP.Show();
                this.Visible = true;
                BEP.UnsetAll();
            }
            else if (Purchase)
            {
                AddPurchase.connString = connString;
                AddPurchase OBJ = new AddPurchase();
                OBJ.Show();
                this.Visible = true;
            }
            else if (Consumption)
            {
                ADDConsumption.connString = connString;
                ADDConsumption UE = new ADDConsumption();
                UE.Show();
                this.Visible = true;
            }
            else if (Pending)
            {

            }
            else if (EquipName)
            {
                AdditionEquipList.connString = connString;
                AdditionEquipList UE = new AdditionEquipList();
                UE.Show();
                this.Visible = true;
            }
            else if (DeptName)
            {
                AddNewDept.connString = connString;
                AddNewDept UE = new AddNewDept();
                UE.Show();
                this.Visible = true;
            }
            else if (Stock)
            {
                ADDStock.connString = connString;
                ADDStock UE = new ADDStock();
                UE.Show();
                this.Visible = true;
                //AddNewDept UE = new AddNewDept();
                //UE.Show();
                //this.Visible = true;
            }
            else if (Stock)
            {
                ADDStock.connString = connString;
                ADDStock UE = new ADDStock();
                UE.Show();
                this.Visible = true;
                //AddNewDept UE = new AddNewDept();
                //UE.Show();
                //this.Visible = true;
            }
            else if (ReminderTasks)
            {

                //Reminder.connString = connString;
                //int i = EquipGrid.SelectedCells[0].RowIndex;
                //Reminder.ItemID = EquipGrid.Rows[i].Cells["Equipment_ID"].Value.ToString();
                //Reminder UE = new Reminder();
                //UE.Show();
                //this.Visible = true;
                Reminder.connString = connString;
                Reminder UE = new Reminder();
                UE.Show();
                this.Visible = true;

            }


        }

        private void btnPurch_Click(object sender, EventArgs e)
        {
            ButtonClick();
            Purchase = true;
            btnPurch.ForeColor = System.Drawing.Color.Red;
            groupBoxPurc.Visible = true;


            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Purchase order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Purchase");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Purchase";
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int i;

            progressBar.Minimum = 0;
            progressBar.Maximum = 200;

            for (i = 0; i <= 200; i++)
            {
                progressBar.Value = i;
            }

        }

        private void cmbIdRepair_MouseDown(object sender, MouseEventArgs e)
        {
            // string ID = cmbIdRepair.Text;
            SqlDataAdapter da;
            SqlCommand cmd;
            using (SqlConnection con = new SqlConnection(connString))
            {
                cmd = new SqlCommand("Select Equipment_ID,Equipment_Name from new_EQUIPMENT_INFO where Equip_status='OK' and Unit_Name= 'Beas Hosp' order by Equipment_ID asc  ", con);

                da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    cmbIdRepair.Items.Add(ds.Tables[0].Rows[i][0]);// + "    " + ds.Tables[0].Rows[i][1]);
                }
                //SqlDataAdapter dataadapter = new SqlDataAdapter(sql, con);
                //DataSet ds = new DataSet();
                //dataadapter.Fill(ds, "new_EQUIPMENT_INFO ");
                //EquipGrid.DataSource = ds;
                //EquipGrid.DataMember = "new_EQUIPMENT_INFO ";
                //EquipGrid.Refresh();
                //con.Close();

            }
        }

        private void cmbIdCons_MouseDown(object sender, MouseEventArgs e)
        {
            SqlCommand cmdcons;
            SqlDataReader dr;
            String hospital = cmbHosp.Text;
            using (SqlConnection con = new SqlConnection(connString))
            {
                cmdcons = new SqlCommand("Select Equipment_ID from new_equipment_info order by Equipment_ID asc", con);

                //  using (SqlCommand cmd = new SqlCommand(cmdString, con))
                {
                    //----adding equipment names to combocon.Open();
                    con.Open();
                    dr = cmdcons.ExecuteReader();
                    while (dr.Read()) //loop
                        cmbIdCons.Items.Add(dr["Equipment_ID"].ToString());
                    con.Close();

                }
            }
        }

        private void maintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ButtonClick();
            Maintenance = true;

            //Purchase = Equipment = Consumption = Repair = Pending =Stock = false;            
            //groupBox3.Visible = false;
            //groupBoxRepair.Visible = false;
            //groupBoxEquip.Visible = false;
            //groupBoxPurc.Visible = false;
            //groupBoxConsumption.Visible = false;
            using (SqlConnection con = new SqlConnection(connString))
            {
                string p = ("Select Equipment_ID, Equipment_Name, Complete_specification,Date_of_purchase,Cost_of_Equipment from New_Equipment_info order by Equipment_ID desc ");
                // select filename,status,max_date = max(dates)from some_table tgroup by filename , status having status = '<your-desired-status-here>'

                // string p = ("select Status, Rdate from Repair_master  group by  having status='Pending'");

                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "New_Equipment_info");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "New_Equipment_info";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                // EquipGrid.Columns.("Phone").Width = 30
                //EquipGrid.Columns[0].Width = 60;
                //EquipGrid.Columns[1].Width = 90;
                //EquipGrid.Columns[2].Width = 100;
                //EquipGrid.Columns[3].Width = 150;
                //EquipGrid.Columns[4].Width = 300;
                //EquipGrid.Columns[5].Width = 100;
                //EquipGrid.Columns[6].Width = 80;
                //EquipGrid.Columns[7].Width = 150;
                //EquipGrid.Columns[8].Width = 100;
                //EquipGrid.Columns[9].Width = 50;

            }
        }

      
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try 
            {
                int month = DateTime.Parse("1." + CmbMonth.Text + "").Month;
                int Year = DateTime.Parse("1." + CmbYear.Text + "").Year;
                string Cat = cmbCategory.Text;
                string p;
                using (SqlConnection con = new SqlConnection(connString))
                {


                    if (Cat.Equals("All"))

                        p = ("select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Equip_status = 'Ok' and new_equipment_info.Unit_name = 'Beas Hosp'and Repair_Master.Repair_maint = 'Maint.'  and MONTH(Rdate) = " + month + "  and   YEAR(Rdate) = " + Year + " order by Rdate asc");
                    else
                        p = ("select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Equip_status = 'Ok' and new_equipment_info.Unit_name = 'Beas Hosp'and new_equipment_info.Category= '" + Cat + "' and Repair_Master.Repair_maint = 'Maint.'  and MONTH(Rdate) = " + month + "  and   YEAR(Rdate) = " + Year + " order by Rdate asc");
                    //  else order by Department desc ");
                    //  if (Dept.Equals("All"))

                    //p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                    //  else 

                    // p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                    con.Open();
                    SqlDataAdapter Sda = new SqlDataAdapter(p, con);
                    DataTable Dt = new DataTable();
                    EquipGrid.DataSource = Dt;
                    Sda.Fill(Dt);
                    con.Close();
                    Maintload();
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());

            }

        }



        private void btnSearchDue_Click(object sender, EventArgs e)
        {
            int month = DateTime.Parse("1." + CmbMonth.Text + "").Month;
            int Year = DateTime.Parse("1." + CmbYear.Text + "").Year;
            string Cat = cmbCategory.Text;
            string p;

            using (SqlConnection con = new SqlConnection(connString))
            {
                if (Cat.Equals("All"))
                    p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS DateAdd from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp'and  Repair_Master.Repair_maint='Maint.' and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");
                else
                    p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS DateAdd from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp'and  Repair_Master.Repair_maint='Maint.' and  new_equipment_info.Category= '" + Cat + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                //  else order by Department desc ");
                //  if (Dept.Equals("All"))

                //p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                //  else 

                // p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                con.Open();
                SqlDataAdapter Sda = new SqlDataAdapter(p, con);
                DataTable Dt = new DataTable();
                EquipGrid.DataSource = Dt;
                Sda.Fill(Dt);
                con.Close();
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Maintload();

            }

        }

        private void deleteDuplicateToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Do you want to Delete Duplicate Record", "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string sql = ("WITH CTE AS (SELECT *, RN = ROW_NUMBER()OVER(PARTITION BY[DOC_NO],[DRUG_SRNO]  ORDER BY DOC_NO desc)FROM[Equipment].[dbo].[consumption]) Delete CTE where RN > 1");

                using (SqlConnection con = new SqlConnection(connString))

                using (SqlCommand deleteRecord = new SqlCommand(sql, con))

                {
                    con.Open();

                    deleteRecord.ExecuteNonQuery();
                    MessageBox.Show("Record Deleted Successfully!");

                }
            }
            else
            {
                MessageBox.Show("Record not Deleted", "Delete Record", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            ButtonClick();
            Consumption = true;

            //Equipment = Repair = Purchase = Maintenance = Pending= Stock = false;            
            //groupBox3.Visible = false;
            //groupBoxConsumption.Visible = true;
            //groupBoxEquip.Visible = false;
            //groupBoxPurc.Visible = false;
            //groupBoxRepair.Visible = false;
            //groupBoxMaint.Visible = false;
            groupBoxConsumption.Location = new Point(300, 10);

            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from consumption  order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "consumption");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "consumption";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Consumpload();
            }


        }

        private void StockView()
        {
            //ButtonClick(); 
            Stock = true;

            //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = false;

            //groupBox3.Visible = false;
            //groupBoxConsumption.Visible = false;
            //groupBoxEquip.Visible = false;
            //groupBoxPurc.Visible = false;
            //groupBoxRepair.Visible = false;
            //groupBoxMaint.Visible = false;
            //groupBoxStock.Visible = true;
            groupBoxStock.Location = new Point(230, 40);

            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Stock");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                //try
                //{
                dataadapter.Fill(ds, "Stock");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Stock";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                //code to count sum
                decimal Total = 0;

                for (int i = 0; i < EquipGrid.Rows.Count; i++)
                {
                    Total += Convert.ToDecimal(EquipGrid.Rows[i].Cells["BAL_VAL"].Value);
                }

                lblAmount.Text = Total.ToString();

                Stockload();
            }
            //catch {
            //    MessageBox.Show("No stck available");

            //};



        }
        private void btnStock_Click(object sender, EventArgs e)
        {
            ButtonClick();
            groupBoxStock.Visible = true;
            btnStock.ForeColor = System.Drawing.Color.Red;
            StockView();
        }

        private void addToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ADDStock.connString = connString;
            ADDStock UE = new ADDStock();
            UE.Show();
            this.Visible = true;
        }



        private void buttonSearchStock_Click_1(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                //select * from bio_db.daily_data2 where Date between '" + dtFrom.ToString("MM/dd/yyyy")+ "' and '" + dtTo.ToString("MM/dd/yyyy") + "' ", mcon)
                //where Date between '" + datefrom.Value.ToString() + "' and '" + dateto.Value.ToString() + "' ", mcon);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Stock  Where DRUG_NM LIKE '%" + textSearchStock.Text + "%'", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Stockload();

            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            StockView();
        }

        private void EquipGrid_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (e.ColumnIndex == 1)
                    //UpdateEquipNew.connString = connString;
                    //int i = EquipGrid.SelectedCells[0].RowIndex;
                    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                UpdateEquipNew UE = new UpdateEquipNew();
                UE.Show();
                this.Visible = true;

            }

            catch
            {

            }
            //StateMents you Want to execute to Get Data 

        }

        private void EquipGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


            ////if (e.ColumnIndex == 10)
            //if (e.ColumnIndex == EquipGrid.Columns["Edit"].Index)

            //    {  //check id value for current row

            //    //if (EquipGrid.Current;Row.Cells[5].Value != null)
            //    UpdateEquipNew.connString = connString;               
            //    int i = EquipGrid.SelectedCells[0].RowIndex;
            //    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
            //    UpdateEquipNew UE = new UpdateEquipNew();
            //    UE.Show();
            //    this.Visible = true;

            //    }

            //else if (e.ColumnIndex == EquipGrid.Columns["View"].Index)

            //{  //check id value for current row

            //    //if (EquipGrid.Current;Row.Cells[5].Value != null)
            //    Equip_Detail_new.connString = connString;
            //    int i = EquipGrid.SelectedCells[0].RowIndex;
            //    Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
            //    Equip_Detail_new UE = new Equip_Detail_new();
            //    UE.Show();
            //    this.Visible = true;

            //}


        }

        private void btnEquiName_Click(object sender, EventArgs e)
        {
            ButtonClick();

            btnEquiName.ForeColor = System.Drawing.Color.Red;
            EquipName = true;


            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Equip_Names  order by Equip_Name asc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Equip_Names");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Equip_Names";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                EquipGrid.Columns[0].Width = 50;
                EquipGrid.Columns[1].Width = 200;
            }
        }
        private void EquipGrid_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                contextMenuStrip1.Show(Cursor.Position.X, Cursor.Position.Y);
            }
        }


        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            // this.Close();
        }

        private void EquipGrid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                try
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        this.EquipGrid.Rows[e.RowIndex].Selected = true;
                        this.rowIndex = e.RowIndex;
                        this.EquipGrid.CurrentCell = this.EquipGrid.Rows[e.RowIndex].Cells[1];
                        this.contextMenuStrip1.Show(this.EquipGrid, e.Location);
                        contextMenuStrip1.Show(Cursor.Position);
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            if (!this.EquipGrid.Rows[this.rowIndex].IsNewRow)
            {
                //if (e.ColumnIndex == EquipGrid.Columns["Edit"].Index)
                ToolStripItem clickedItem = sender as ToolStripItem;
                // your code here
                {  //check id value for current row

                    //if (EquipGrid.Current;Row.Cells[5].Value != null)
                    UpdateEquipNew.connString = connString;

                    UpdateEquipNew.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    UpdateEquipNew UE = new UpdateEquipNew();
                    UE.Show();
                    this.Visible = true;

                    //    } //this.EquipGrid.Rows.RemoveAt(this.rowIndex);
                }
            }
        }

        private void viewToolStripMenuItem4_Click(object sender, EventArgs e)
        {

            if (Equipment)
            {
                try

                {
                    Equip_Detail_new.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    //AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                    Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_Id"].Value.ToString();
                    Equip_Detail_new UE = new Equip_Detail_new();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }


            }
            else if (Repair)
            {
                Equip_Detail_new.connString = connString;

                Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                Equip_Detail_new UE = new Equip_Detail_new();
                UE.Show();
                this.Visible = true;
                

            }
            else if (Purchase)
            {
            }
            else if (Consumption)
            {
                Equip_Detail_new.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["eqid"].Value.ToString();
                Equip_Detail_new UE = new Equip_Detail_new();

               
                //EditConsumption.ItemID = EquipGrid.Rows[i].Cells["eqid"].Value.ToString();
                //EditConsumption.ID = EquipGrid.Rows[i].Cells["ID"].Value.ToString();
                //EditConsumption UE = new EditConsumption();
                UE.Show();
                this.Visible = true;

            }
            else if (Maintenance)
            {
                try

                {
                    Equip_Detail_new.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                    Equip_Detail_new UE = new Equip_Detail_new();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }

            }
            else if (Pending)
            {
                {
                    Equip_Detail_new.connString = connString;

                    Equip_Detail_new.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                    Equip_Detail_new UE = new Equip_Detail_new();
                    UE.Show();
                    this.Visible = true;

                }

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void repairToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                try
                {

                    RepairView.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    RepairView.equipid = this.EquipGrid.CurrentRow.Cells["Equipment_ID"].Value.ToString();
                    RepairView.Dept = this.EquipGrid.CurrentRow.Cells["Department"].Value.ToString();
                    RepairView UE = new RepairView();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }

            }
            else if (Repair)
            {
                try
                {

                    RepairView.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    RepairView.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                    RepairView.Dept = this.EquipGrid.CurrentRow.Cells["Dept"].Value.ToString();
                    RepairView UE = new RepairView();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }
            }
            else if (Maintenance)
            {
                try
                {

                    RepairView.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    RepairView.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                    RepairView.Dept = this.EquipGrid.CurrentRow.Cells["Dept"].Value.ToString();
                    RepairView UE = new RepairView();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }

            }
            else if (Purchase)
            {

            }
            else if (Consumption)
            {

            }
            else if (Pending)
            {
                try
                {

                    RepairView.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    RepairView.equipid = this.EquipGrid.CurrentRow.Cells["Eq_ID"].Value.ToString();
                    RepairView.Dept = this.EquipGrid.CurrentRow.Cells["Dept"].Value.ToString();
                    RepairView UE = new RepairView();
                    UE.Show();
                    this.Visible = true;
                }

                catch
                {

                }
            }
        }

        private void maintToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                try
                {
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    WDV_New.Cost = this.EquipGrid.CurrentRow.Cells["Cost_Of_Equipment"].Value.ToString();
                    WDV_New.purchdate = this.EquipGrid.CurrentRow.Cells["Date_Of_Purchase"].Value.ToString();
                    WDV_New BE = new WDV_New();
                    BE.Show();
                    this.Visible = true;
                    //lblID.Text = equipid;

                    //RepairView UE = new RepairView();

                    //WDV_New.Cost = lblPrice.Text;
                    //WDV_New.purchdate = lblDOP.Text;
                    //WDV_New BE = new WDV_New();
                    //BE.Show();
                    //this.Visible = true;
                    //lblID.Text = equipid;
                }

                catch
                {

                }

            }
            else if (Repair)
            {


            }
            else if (Purchase)
            {

            }
            else if (Consumption)
            {

            }
            else if (Pending)
            {
                AddRepair.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();

                AddRepair VE = new AddRepair("Repair");
                VE.Show();
                this.Visible = true;
            }

        }

        private void btnReminder_Click(object sender, EventArgs e)
        {
            ButtonClick();
            ReminderTasks = true;
            btnReminder.ForeColor = System.Drawing.Color.Red;

            ReminderLoad();
           //ReminderCheck();



        }
        private void ReminderLoad()
        {
            ReminderTasks = true;
            btnReminder.ForeColor = System.Drawing.Color.Red;


            using (SqlConnection con = new SqlConnection(connString))
            {
                con.Open();
                SqlDataAdapter dataadapter = new SqlDataAdapter("Select * from Reminders", con);
                DataSet ds = new DataSet();
                DataTable tbl = new DataTable();
                dataadapter.Fill(tbl);
                EquipGrid.DataSource = tbl;

                con.Close();
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;

            }
        }
        private void ReminderCheck()
        {
            DateTime One, Two = DateTime.Now;
            //TimeSpan Differences;
            //int Diff2 = -1;            
            string Q = "select * from reminders where ISReminderSet='True'";
            using (SqlConnection con = new SqlConnection(connString))
            {
                con.Open();
                SqlCommand CM = new SqlCommand(Q);
                CM.Connection = con;
                SqlDataReader Rd = CM.ExecuteReader();
                if (Rd.HasRows)
                {
                    while (Rd.Read())
                    {
                        // Desc = Rd[2].ToString();
                        One = Convert.ToDateTime(Rd["DueDate"].ToString());
                        //Differences = One - Two;
                        //Diff2 = int.Parse(Differences.Days.ToString()) - int.Parse(Rd["Reminder_Before_days"].ToString());
                        Two = DateTime.UtcNow.Date;
                        string StaffName = (Rd["Staff_Name"].ToString());
                        string Description = (Rd["Description"].ToString());
                        string Task = (Rd["ActionRequired"].ToString());
                        bool R = Boolean.Parse(Rd["Recuring"].ToString());
                        int T = Convert.ToInt32(Rd["Reccuring_Time"].ToString());
                        //string ActionBy = (Rd["Staff_Name"].ToString());
                        id = Convert.ToInt32(Rd["ID"].ToString());
                        int totalDaysLeft = Convert.ToInt32((One.Date - DateTime.UtcNow.Date).TotalDays);
                        int daysBefore = Convert.ToInt32(Rd["Reminder_Before_days"].ToString());
                        int reminderday = totalDaysLeft - daysBefore;

                        if ((reminderday <= 0) && (One >= Two)) /*|| totalDaysLeft <= daysBefore )*/
                        {
                            ReminderPopUp.Descrition = Description;
                            ReminderPopUp.Task = Task;
                            ReminderPopUp.StaffName = StaffName;
                            ReminderPopUp.DayRemaining = totalDaysLeft.ToString();
                            ReminderPopUp.Duedate = One.ToString("dd-MM-yyyy");


                            ReminderPopUp Obj = new ReminderPopUp();
                            Obj.Show();
                            if ((One <= Two) && R == true)
                            {
                                SqlConnection CONN = new SqlConnection(connString);
                                CONN.Open();
                                //var XY = Convert.ToDateTime(One.AddDays(T));
                                //UPDATE classes  SET `date` = DATE_ADD(`date`, INTERVAL 2 DAY) DATEADD(day, 2, DepartureDate)
                                String Q1 = "Update reminders set duedate = DATEADD(day," + T + ", duedate) where id = " + id;
                                SqlCommand UpQuery = new SqlCommand(Q1);
                                UpQuery.Connection = CONN;
                                UpQuery.ExecuteNonQuery();
                                //MessageBox.Show("Recurring Reminder done!!");
                                CONN.Close();

                            }

                        }
                        else
                        {
                            //MessageBox.Show("No active reminder for today!!");
                            //return;
                        }
                    }
                }
                   
                
                con.Close();
            }
        }

        private void addTaskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                try
                {
                    Reminder.connString = connString;
                    int i = EquipGrid.SelectedCells[0].RowIndex;
                    Reminder.EquipID = EquipGrid.Rows[i].Cells["Equipment_ID"].Value.ToString();
                    Reminder UE = new Reminder();
                    UE.Show();
                    this.Visible = true;
                    //Reminder.connString = connString;
                    //int i = EquipGrid.SelectedCells[0].RowIndex;
                    //Reminder.ID = EquipGrid.Rows[i].Cells["Equipment_ID"].Value.ToString();
                    //Reminder UE = new Reminder();
                    //UE.Show();
                    //this.Visible = true;

                }

                catch
                {

                }

            }
            else if (Repair)
            {
                //AddRepair.connString = connString;
                //int i = EquipGrid.SelectedCells[0].RowIndex;
                //AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                //AddRepair VE = new AddRepair("Repair");
                //VE.Show();
                //this.Visible = true;

            }
            else if (Maintenance)
            {
                //AddRepair.connString = connString;
                //int i = EquipGrid.SelectedCells[0].RowIndex;
                //AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                //AddRepair VE = new AddRepair("Repair");
                //VE.Show();
                //this.Visible = true;

            }
            else if (Pending)
            {
                //AddRepair.connString = connString;
                //int i = EquipGrid.SelectedCells[0].RowIndex;
                //AddRepair.rid = EquipGrid.Rows[i].Cells[0].Value.ToString();
                //AddRepair VE = new AddRepair("Repair");
                //VE.Show();
                //this.Visible = true;

            }
            else if (EquipName)
            {
                //AdditionEquipList.connString = connString;
                //AdditionEquipList UE = new AdditionEquipList();
                //UE.Show();
                //this.Visible = true;

            }

            else if (DeptName)
            {
                //AddNewDept.connString = connString;
                //AddNewDept UE = new AddNewDept();
                //UE.Show();
                //this.Visible = true;

            }
            else if (Consumption)
            {
                //EditConsumption.connString = connString;
                //int i = EquipGrid.SelectedCells[0].RowIndex;
                //EditConsumption.ItemID = EquipGrid.Rows[i].Cells[3].Value.ToString();
                //EditConsumption UE = new EditConsumption();
                //UE.Show();
                //this.Visible = true;

            }

            else if (ReminderTasks)
            {
                Reminder.connString = connString;
                int i = EquipGrid.SelectedCells[0].RowIndex;
                Reminder.ID = EquipGrid.Rows[i].Cells[3].Value.ToString();
                Reminder UE = new Reminder();
                UE.Show();
                this.Visible = true;

            }

        }

        private void btnDepName_Click(object sender, EventArgs e)
        {

            ButtonClick();          
            btnDepName.ForeColor = System.Drawing.Color.Red;           
            DeptName = true;

            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from Dept_Details  order by Department asc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Dept_Details");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Dept_Details";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                EquipGrid.Columns[0].Width = 50;
                EquipGrid.Columns[1].Width = 200;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
        }

        private void timerReminder_Tick(object sender, EventArgs e)
        {

            //ReminderPopUp.connString = connString;
            lbltimecheck.Text = (int.Parse(lbltimecheck.Text) - 1).ToString(); //lowering the value - explained above
            if (int.Parse(lbltimecheck.Text) == 0) //if the countdown reaches '0', we stop it   
            {
                timerReminder.Stop();
                ReminderCheck();
            }
                //    DateTime One, Two = DateTime.Now;
                //    //TimeSpan Differences;
                //    //int Diff2 = -1;            
                //    string Q = "select * from reminders where ISReminderSet='True'";
                //    using (SqlConnection con = new SqlConnection(connString))
                //    {
                //        con.Open();
                //        SqlCommand CM = new SqlCommand(Q);
                //        CM.Connection = con;
                //        SqlDataReader Rd = CM.ExecuteReader();
                //        if (Rd.HasRows)
                //        {
                //            while (Rd.Read())
                //            {
                //                // Desc = Rd[2].ToString();
                //                 One = Convert.ToDateTime(Rd["DueDate"].ToString());
                //                //Differences = One - Two;
                //                //Diff2 = int.Parse(Differences.Days.ToString()) - int.Parse(Rd["Reminder_Before_days"].ToString());
                //                Two = DateTime.UtcNow.Date;
                //                string StaffName = (Rd["Staff_Name"].ToString());
                //                string Description = (Rd["Description"].ToString());
                //                string Task = (Rd["ActionRequired"].ToString());
                //                bool R = Boolean.Parse(Rd["Recuring"].ToString());
                //                int T = Convert.ToInt32(Rd["Reccuring_Time"].ToString());
                //                //string ActionBy = (Rd["Staff_Name"].ToString());
                //                id = Convert.ToInt32(Rd["ID"].ToString());
                //                int totalDaysLeft = Convert.ToInt32((One.Date - DateTime.UtcNow.Date).TotalDays);
                //                int daysBefore = Convert.ToInt32(Rd["Reminder_Before_days"].ToString());
                //                int reminderday =totalDaysLeft-daysBefore ;

                //                if ((reminderday == 0) || (One == Two)) /*|| totalDaysLeft <= daysBefore )*/
                //                {
                //                    ReminderPopUp.Descrition = Description;
                //                    ReminderPopUp.Task = Task;
                //                    ReminderPopUp.StaffName = StaffName;

                //                    ReminderPopUp Obj = new ReminderPopUp();
                //                    Obj.Show();
                //                    if((reminderday ==0) && R == true)
                //                    {
                //                        SqlConnection CONN = new SqlConnection(connString);
                //                        CONN.Open();
                //                        //var XY = Convert.ToDateTime(One.AddDays(T));
                //                        //UPDATE classes  SET `date` = DATE_ADD(`date`, INTERVAL 2 DAY) DATEADD(day, 2, DepartureDate)
                //                        String Q1 = "Update reminders set duedate = DATEADD(day," + T + ", duedate) where id = "+id;
                //                        SqlCommand UpQuery = new SqlCommand(Q1);
                //                        UpQuery.Connection = CONN;
                //                        UpQuery.ExecuteNonQuery();
                //                        //MessageBox.Show("Recurring Reminder done!!");
                //                        CONN.Close();

                //                    }

                //                }
                //                else
                //                {
                //                    //MessageBox.Show("No record");
                //                }


                //            }
                //            return;
                //        }

                //        con.Close();




                //}

                //    //DateTime Date1 = Convert.ToDateTime(dtReminder.Value);
                //    //DateTime Date2 = DateTime.Now;
                //    //TimeSpan sD = Date1 - Date2;
                //    //int XX = int.Parse(sD.Days.ToString()) - int.Parse(cmbDays.SelectedItem.ToString());
                //    //DaysLeft.Text = "Days Left for Reminder " + XX;
            }

        private void addToolStripMenuItem3_Click(object sender, EventArgs e)
        {

            Add_New_Staff_Member.connString = connString;
            Add_New_Staff_Member VE = new Add_New_Staff_Member();
            VE.Show();
            this.Visible = true;
        }

        private void viewToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            ButtonClick();
            groupBoxStock.Visible = true;
            btnStock.ForeColor = System.Drawing.Color.Red;
            StockView();
        }

        private void viewToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            Staff.connString = connString;
            Staff VE = new Staff();
            VE.Show();
            this.Visible = true;
        }

        private void btnEquipSearch_Click(object sender, EventArgs e)
        {
            EquipSearch();

        }
        private void EquipSearch()
        {

            using (SqlConnection con = new SqlConnection(connString))
            {
                string p = ("Select ID, EQUIPMENT_ID,EQUIPMENT_NAME,COMPLETE_SPECIFICATION,DEPARTMENT,DATE_OF_PURCHASE,COST_OF_EQUIPMENT,Category,EQUIP_STATUS,Unit_Name From new_EQUIPMENT_INFO ");


                if (cmbHosp.Text.Equals("All")) /*&& equip.Equals("All") && hospital.Equals("All") && status.Equals("All"))*///0000
                {
                    sql = p + "where EQUIP_STATUS= 'OK' and CONCAT([Complete_Specification], [Equipment_Name],[EQUIPMENT_ID], DATEPART(yyyy,DATE_OF_PURCHASE)) LIKE '%" + txtSearch.Text + "%' ";
                }

                //CONCAT([First Name], [Last Name]) // Convert.ToDateTime(dateTimePicker1.Text) CONVERT(varchar, DateFieldName, 101) 
                else
                {
                    sql = p + " where Unit_Name= '" + cmbHosp.Text + "' and EQUIP_STATUS= 'OK' and CONCAT([Complete_Specification],[EQUIPMENT_ID], [Equipment_Name], DATEPART(yyyy, DATE_OF_PURCHASE))  LIKE '%" + txtSearch.Text +  "%'";
                    con.Open();
                }

       

                SqlDataAdapter dataadapter = new SqlDataAdapter(sql, con);
                DataSet ds = new DataSet();
                DataTable tbl = new DataTable();
                dataadapter.Fill(tbl);
                EquipGrid.DataSource = tbl;

                con.Close();
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;



            }

        }

        private void btnExportXls_Click(object sender, EventArgs e)
        {
            string fileName;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog1.Title = "To Excel";
            saveFileDialog1.FileName = this.Text + " (" + DateTime.Now.ToString("yyyy-MM-dd") + ")";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(this.Text);
                for (int i = 0; i < EquipGrid.Columns.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = EquipGrid.Columns[i].Name;
                }

                for (int i = 0; i < EquipGrid.Rows.Count; i++)
                {
                    for (int j = 0; j < EquipGrid.Columns.Count; j++)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = EquipGrid.Rows[i].Cells[j].Value.ToString();

                        if (worksheet.Cell(i + 2, j + 1).Value.ToString().Length > 0)
                        {
                            XLAlignmentHorizontalValues align;

                            switch (EquipGrid.Rows[i].Cells[j].Style.Alignment)
                            {
                                case DataGridViewContentAlignment.BottomRight:
                                    align = XLAlignmentHorizontalValues.Right;
                                    break;
                                case DataGridViewContentAlignment.MiddleRight:
                                    align = XLAlignmentHorizontalValues.Right;
                                    break;
                                case DataGridViewContentAlignment.TopRight:
                                    align = XLAlignmentHorizontalValues.Right;
                                    break;

                                case DataGridViewContentAlignment.BottomCenter:
                                    align = XLAlignmentHorizontalValues.Center;
                                    break;
                                case DataGridViewContentAlignment.MiddleCenter:
                                    align = XLAlignmentHorizontalValues.Center;
                                    break;
                                case DataGridViewContentAlignment.TopCenter:
                                    align = XLAlignmentHorizontalValues.Center;
                                    break;

                                default:
                                    align = XLAlignmentHorizontalValues.Left;
                                    break;
                            }

                            worksheet.Cell(i + 2, j + 1).Style.Alignment.Horizontal = align;

                            XLColor xlColor = XLColor.FromColor(EquipGrid.Rows[i].Cells[j].Style.SelectionBackColor);
                            worksheet.Cell(i + 2, j + 1).AddConditionalFormat().WhenLessThan(1).Fill.SetBackgroundColor(xlColor);

                            worksheet.Cell(i + 2, j + 1).Style.Font.FontName = EquipGrid.Font.Name;
                            worksheet.Cell(i + 2, j + 1).Style.Font.FontSize = EquipGrid.Font.Size;

                        }
                    }
                }
                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(fileName);
                MessageBox.Show("Data is exported to EXLS successfully");
            }
        }
    

        public void ExportToExcelWithFormatting(DataGridView EquipGrid)
         {
           string fileName;

          SaveFileDialog saveFileDialog1 = new SaveFileDialog();
          saveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog1.Title = "To Excel";
                saveFileDialog1.FileName = this.Text + " (" + DateTime.Now.ToString("yyyy-MM-dd") + ")";

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add(this.Text);
                    for (int i = 0; i < EquipGrid.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = EquipGrid.Columns[i].Name;
                    }

                    for (int i = 0; i < EquipGrid.Rows.Count; i++)
                    {
                        for (int j = 0; j < EquipGrid.Columns.Count; j++)
                        {
                            worksheet.Cell(i + 2, j + 1).Value = EquipGrid.Rows[i].Cells[j].Value.ToString();

                            if (worksheet.Cell(i + 2, j + 1).Value.ToString().Length > 0)
                            {
                                XLAlignmentHorizontalValues align;

                                switch (EquipGrid.Rows[i].Cells[j].Style.Alignment)
                                {
                                    case DataGridViewContentAlignment.BottomRight:
                                        align = XLAlignmentHorizontalValues.Right;
                                        break;
                                    case DataGridViewContentAlignment.MiddleRight:
                                        align = XLAlignmentHorizontalValues.Right;
                                        break;
                                    case DataGridViewContentAlignment.TopRight:
                                        align = XLAlignmentHorizontalValues.Right;
                                        break;

                                    case DataGridViewContentAlignment.BottomCenter:
                                        align = XLAlignmentHorizontalValues.Center;
                                        break;
                                    case DataGridViewContentAlignment.MiddleCenter:
                                        align = XLAlignmentHorizontalValues.Center;
                                        break;
                                    case DataGridViewContentAlignment.TopCenter:
                                        align = XLAlignmentHorizontalValues.Center;
                                        break;

                                    default:
                                        align = XLAlignmentHorizontalValues.Left;
                                        break;
                                }

                                worksheet.Cell(i + 2, j + 1).Style.Alignment.Horizontal = align;

                                XLColor xlColor = XLColor.FromColor(EquipGrid.Rows[i].Cells[j].Style.SelectionBackColor);
                                worksheet.Cell(i + 2, j + 1).AddConditionalFormat().WhenLessThan(1).Fill.SetBackgroundColor(xlColor);

                                worksheet.Cell(i + 2, j + 1).Style.Font.FontName = EquipGrid.Font.Name;
                                worksheet.Cell(i + 2, j + 1).Style.Font.FontSize = EquipGrid.Font.Size;

                            }
                        }
                    }
                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(fileName);
                    MessageBox.Show("Data is exported to EXLS successfully");
                }
               
            }

        private void deleteToolStripMenuItem3_Click(object sender, EventArgs e)
        {

            if (Equipment)
            {
                try

                {
                    
                }

                catch
                {

                }


            }
            else if (Repair)
            {
               

            }
            else if (Purchase)
            {
            }
            else if (Consumption)
            {

            }
            else if (Maintenance)
            {
                try

                {
                  
                }

                catch
                {

                }

            }
            else if (Pending)
            {
                try

                {

                }

                catch
                {

                }

            }
            else if (ReminderTasks)
            {
                try

                {
                    if (MessageBox.Show("you want to Delete This Record", "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {

                        // No row selected no delete....
                        if (EquipGrid.SelectedRows.Count == 0)
                        {
                            MessageBox.Show("No row selected !");// show a message here to inform
                        }

                        // Prepare the command text with the parameter placeholder

                        string sql = "DELETE FROM  Reminders  WHERE ID=@rowID";
                        using (SqlConnection con = new SqlConnection(connString))

                        // Create the connection and the command inside a using block
                        //using (SqlConnection myConnection = new SqlConnection("Data Source=Epic-LaptopWR;Initial Catalog=words;Integrated Security=True"))
                        using (SqlCommand deleteRecord = new SqlCommand(sql, con))

                        {
                            con.Open();

                            // if (EquipGrid.CurrentCell.RowIndex >0)
                            {
                                int rowIndex = EquipGrid.CurrentCell.RowIndex;
                                int selectedIndex = EquipGrid.SelectedRows[0].Index;
                                // gets the RowID from the first column in the grid
                                string rowID = Convert.ToString(EquipGrid[0, selectedIndex].Value);

                                deleteRecord.Parameters.AddWithValue("@rowID", rowID);//.Value.ToString();//= rowID;
                                deleteRecord.ExecuteNonQuery();
                                MessageBox.Show("Record Deleted Successfully");

                                // Remove the row from the grid
                                EquipGrid.Rows.RemoveAt(selectedIndex);
                                string qty = Convert.ToString(EquipGrid.Rows.Count);
                                lblCount.Text = qty;
                                //EquipGrid.Rows.RemoveAt(EquipGrid.CurrentRow.Index);
                            }

                        }// end of delete Button 
                    }
                    else
                    {
                        MessageBox.Show("Record not Deleted", "Delete Record", MessageBoxButtons.OK, MessageBoxIcon.Information);
                      


                    }

                }

                catch
                {

                }

            }
        }

        private void btnRefreshGrid_Click(object sender, EventArgs e)
        {
            if (Equipment)
            {
                showData();
            }
            else if (Repair)
            {
                ButtonClick();
                //btnRepair.Text = "Repair >>";
                //btnEquip.TextAlign = ContentAlignment.MiddleRight;
                btnRepair.ForeColor = System.Drawing.Color.Red;
                groupBoxRepair.Visible = true;

                //EquipGrid.DataSource = null;
                //EquipGrid.Refresh();
                //EquipGrid.Columns.Clear();
                //EquipGrid.Rows.Clear();


                //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = Reminder=false;
                Repair = true;



                using (SqlConnection con = new SqlConnection(connString))
                {
                    //SELECT Orders.OrderID, Customers.CustomerName FROM Orders iNNER JOIN Customers ON Orders.CustomerID = Customers.CustomerID;
                    string p = ("Select TOP 500 Repair_Master.RID,Repair_Master.Rdate,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id order by Rdate desc ");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataTable Dt = new DataTable();
                    dataadapter.Fill(Dt);
                    //EquipGrid.Refresh();
                    //EquipGrid.Rows.Clear();              

                    EquipGrid.DataSource = Dt;
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    Repairload();

                }
            }
            else if (Purchase)
            {
                ButtonClick();
                Purchase = true;
                btnPurch.ForeColor = System.Drawing.Color.Red;
                groupBoxPurc.Visible = true;


                using (SqlConnection con = new SqlConnection(connString))
                {

                    string p = ("Select * from Purchase order by DOC_DT desc");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds, "Purchase");
                    EquipGrid.DataSource = ds;
                    EquipGrid.DataMember = "Purchase";
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    Purchload();
                }
            }
            else if (Consumption)
            {
                ButtonClick();

                btnCons.ForeColor = System.Drawing.Color.Red;

                Consumption = true;
                groupBoxConsumption.Visible = true;


                using (SqlConnection con = new SqlConnection(connString))
                {

                    string p = ("Select * from consumption  order by DOC_DT desc");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds, "consumption");
                    EquipGrid.DataSource = ds;
                    EquipGrid.DataMember = "consumption";
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    Consumpload();
                }
            }
            else if (Pending)
            {
                ButtonClick();

                btnPend.ForeColor = System.Drawing.Color.Red;
                Pending = true;


                //Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = false;


                using (SqlConnection con = new SqlConnection(connString))
                {
                    //string p = ("Select Repair_Master.RID,Repair_Master.Rdate,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Action_taken,Repair_Master.Status,Repair_Master.Att_by from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where Repair_Master.status='Pending' order by Rdate desc ");
                    string p = @"dbo.[sp_pending]";

                    SqlCommand cmd = new SqlCommand(p, con);
                    /****** Script for SelectTopNRows command from SSMS  ******/

                    // select filename,status,max_date = max(dates)from some_table tgroup by filename , status having status = '<your-desired-status-here>'

                    // string p = ("select Status, Rdate from Repair_master  group by  having status='Pending'");

                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds, "Repair_master");
                    EquipGrid.DataSource = ds;
                    EquipGrid.DataMember = "Repair_master";
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    // EquipGrid.Columns.("Phone").Width = 30
                    EquipGrid.Columns[0].Width = 60;
                    EquipGrid.Columns[1].Width = 90;
                    EquipGrid.Columns[2].Width = 100;
                    EquipGrid.Columns[3].Width = 150;
                    EquipGrid.Columns[4].Width = 300;
                    EquipGrid.Columns[5].Width = 100;
                    EquipGrid.Columns[6].Width = 80;
                    EquipGrid.Columns[7].Width = 350;
                    EquipGrid.Columns[8].Width = 100;
                    EquipGrid.Columns[9].Width = 50;
                    //EquipGrid.Columns[10].Width = 200;

                }
            }
            else if (EquipName)
            {
                ButtonClick();

                btnEquiName.ForeColor = System.Drawing.Color.Red;
                EquipName = true;


                using (SqlConnection con = new SqlConnection(connString))
                {

                    string p = ("Select * from Equip_Names  order by Equip_Name asc");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds, "Equip_Names");
                    EquipGrid.DataSource = ds;
                    EquipGrid.DataMember = "Equip_Names";
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    EquipGrid.Columns[0].Width = 50;
                    EquipGrid.Columns[1].Width = 200;
                }
            }
            else if (DeptName)
            {
                ButtonClick();
                btnDepName.ForeColor = System.Drawing.Color.Red;
                DeptName = true;

                using (SqlConnection con = new SqlConnection(connString))
                {

                    string p = ("Select * from Dept_Details  order by Department asc");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds, "Dept_Details");
                    EquipGrid.DataSource = ds;
                    EquipGrid.DataMember = "Dept_Details";
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    EquipGrid.Columns[0].Width = 50;
                    EquipGrid.Columns[1].Width = 200;
                }
            }
            else if (Stock)
            {
                ButtonClick();
                groupBoxStock.Visible = true;
                btnStock.ForeColor = System.Drawing.Color.Red;
                StockView();
            }
            else if (Stock)
            {
                ADDStock.connString = connString;
                ADDStock UE = new ADDStock();
                UE.Show();
                this.Visible = true;
                //AddNewDept UE = new AddNewDept();
                //UE.Show();
                //this.Visible = true;
            }
            else if (ReminderTasks)
            {

                ButtonClick();
                ReminderTasks = true;
                btnReminder.ForeColor = System.Drawing.Color.Red;

                ReminderLoad();
            }

            else if (Maintenance)
            {

                ButtonClick();
                btnMaint.ForeColor = System.Drawing.Color.Red;
                Maintenance = true;
                groupBoxMaint.Visible = true;
                string Cat = cmbCategory.Text;
                SqlCommand cmdCat;
                SqlDataReader dr;
                //Convert combobox Month Name to number
                //int month = DateTime.Parse("1." + CmbMonth.Text + "").Month;
                //int Year = DateTime.Parse("1." + CmbYear.Text + "").Year;
                // string Category = cmbCategory.Text;

                using (SqlConnection con = new SqlConnection(connString))
                {
                    cmdCat = new SqlCommand("Select Category from Category order  by Category  asc", con);

                    string p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name = 'Beas Hosp'and Repair_Master.Repair_maint = 'Maint.' order by Rdate desc ");
                    // string p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,new_equipment_info.Category,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS DateAdd from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp'and  Repair_Master.Repair_maint='Maint.' order by Department desc ");

                    //  else order by Department desc ");
                    //  if (Dept.Equals("All"))

                    //p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                    //  else 

                    // p = ("Select Repair_Master.RID,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Rdate, DATEADD(Month, 3, Rdate) AS Maint_Due_Date ,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master  iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id where new_equipment_info.Unit_name= 'Beas Hosp' and  new_equipment_info.Department= '" + Dept + "'and Repair_Master.Repair_maint='Maint.'and  MONTH (DATEADD(Month, 3, Rdate) ) = '" + month + "'and   YEAR (DATEADD(Month, 3, Rdate) ) = '" + Year + "' order by Rdate desc ");

                    con.Open();
                    SqlDataAdapter Sda = new SqlDataAdapter(p, con);
                    DataTable Dt = new DataTable();
                    EquipGrid.DataSource = Dt;
                    Sda.Fill(Dt);
                    con.Close();
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;


                    con.Open();
                    dr = cmdCat.ExecuteReader();
                    cmbCategory.Items.Add("All");
                    while (dr.Read()) //loop
                        cmbCategory.Items.Add(dr["Category"].ToString());
                    con.Close();
                    Maintload();
                    try
                    {
                        //clear all data from combobox
                        CmbMonth.Items.Clear();
                        //add default item
                        CmbMonth.Items.Add("Select");
                        //fill array from System.Globalization.DateTimeFormatInfo.InvariantInfo
                        var Months = System.Globalization.DateTimeFormatInfo.InvariantInfo.MonthNames;
                        //loop array for add items
                        foreach (string sm in Months)
                        {
                            if (sm != "")
                                CmbMonth.Items.Add(sm);
                        }
                        //set selected item for display on startup
                        CmbMonth.Text = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                    //Add year in Combobox

                    try
                    {
                        //clear all data from combobox
                        CmbYear.Items.Clear();
                        //add default item
                        CmbYear.Items.Add("Select");
                        //loop array for add items
                        for (int i = DateTime.Now.Year; i < DateTime.Now.Year + 15; i++)
                        {
                            CmbYear.Items.Add(i);
                        }
                        //set selected item for display on startup
                        CmbYear.Text = DateTime.Now.Year.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());

                    }
                }
            }


        }

        //private void cmbIdPurc_MouseDown(object sender, MouseEventArgs e)
        //{
        //    SqlCommand cmdcons;
        //    SqlDataReader dr;
        //    String hospital = cmbHosp.Text;
        //    using (SqlConnection con = new SqlConnection(connString))
        //    {
        //        cmdcons = new SqlCommand("Select Equipment_ID from new_equipment_info order by Equipment_ID asc", con);

        //        //  using (SqlCommand cmd = new SqlCommand(cmdString, con))
        //        {
        //            //----adding equipment names to combocon.Open();
        //            con.Open();
        //            dr = cmdcons.ExecuteReader();
        //            while (dr.Read()) //loop
        //                cmbIdPurc.Items.Add(dr["Equipment_ID"].ToString());
        //            con.Close();

        //        }
        //    }
        //}

        private void btnSearchDate_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                // SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DRUG_NM  LIKE '%" + txtSearchPurc.Text + "%'", con);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DOC_DT between '" + dateTimePickerPurFrom.Value.ToString("yyyy/MM/dd") + "' and '" + dateTimePickerPurTo.Value.ToString("yyyy/MM/dd") + "'", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();

            }
        }

        private void btnSearchDate_Click_1(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                // SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DRUG_NM  LIKE '%" + txtSearchPurc.Text + "%'", con);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM Purchase Where DOC_DT between '" + dateTimePickerPurFrom.Value.ToString("yyyy/MM/dd") + "' and '" + dateTimePickerPurTo.Value.ToString("yyyy/MM/dd") + "'order by DOC_DT desc", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Purchload();

                /*


                BindingSource bs = new BindingSource();
                bs.DataSource = EquipGrid.DataSource;
               // bs.Filter = "DRUG_NM like '%" + txtSearchPurc.Text + "%'";
                EquipGrid.DataSource = bs;
                (bs.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '%{0}%'", txtSearchPurc.Text);
                */
            }
        }

        private void btnSearchDateCons_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection(connString))
            {
                //select * from bio_db.daily_data2 where Date between '" + dtFrom.ToString("MM/dd/yyyy")+ "' and '" + dtTo.ToString("MM/dd/yyyy") + "' ", mcon)
                //where Date between '" + datefrom.Value.ToString() + "' and '" + dateto.Value.ToString() + "' ", mcon);
                SqlDataAdapter Sda = new SqlDataAdapter("SELECT* FROM consumption Where DOC_DT between '" + dateTimePickerConFrom.Value.ToString("yyyy/MM/dd") + "' and '" + dateTimePickerConTo.Value.ToString("yyyy/MM/dd") + "'  order by DOC_DT desc", con);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                EquipGrid.DataSource = Dt;
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Consumpload();

            }
        }

        private void btnOverDue_Click(object sender, EventArgs e)
        {
            //string MaintDate;
            string p;
            string Cat = cmbCategory.Text;
            using (SqlConnection con = new SqlConnection(connString))
            {
                if (Cat.Equals("All"))
                {
                    p = "SELECT t1.[RID],t1.[RDATE] as Max_Date,t1.[EQ_ID],new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification," +
                      "t1.[DEPT],t1.[REPAIR_MAINT],t1.[ACTION_TAKEN],DATEDIFF(day, RDATE, GETDATE()) AS DaysOverdue FROM([Equipment].[dbo].[REPAIR_MASTER] " +
                      "t1 inner join[Equipment].[dbo].new_EQUIPMENT_INFO ON " +
                      "t1.EQ_ID = New_Equipment_info.Equipment_Id inner join (select Eq_ID, max(RDate) as MaxDate FROM " +
                      "[Equipment].[dbo].[REPAIR_MASTER] group by Eq_ID) tm on t1.EQ_ID=tm.EQ_ID and t1.RDATE =tm.MaxDate)" +
                      "where t1.Repair_Maint='maint.'  and DATEDIFF(day, RDATE, GETDATE()) >=" + int.Parse(cmbDueDays.Text) + " and "+
                      "new_equipment_info.Equip_status = 'Ok' and new_equipment_info.Unit_name = 'Beas Hosp'  ORDER BY t1.RDATE desc";

                }
                else
             
                {
                   p = "SELECT t1.[RID],t1.[RDATE] as Max_Date,t1.[EQ_ID],new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification," +
                       "t1.[DEPT],t1.[REPAIR_MAINT],t1.[ACTION_TAKEN],DATEDIFF(day, RDATE, GETDATE()) AS DaysOverdue FROM([Equipment].[dbo].[REPAIR_MASTER] " +
                       "t1 inner join[Equipment].[dbo].new_EQUIPMENT_INFO ON " +
                       "t1.EQ_ID = New_Equipment_info.Equipment_Id inner join (select Eq_ID, max(RDate) as MaxDate FROM " +
                       "[Equipment].[dbo].[REPAIR_MASTER] group by Eq_ID) tm on t1.EQ_ID=tm.EQ_ID and t1.RDATE =tm.MaxDate)" +
                       "where t1.Repair_Maint='maint.' and DATEDIFF(day, RDATE, GETDATE()) >" + int.Parse(cmbDueDays.Text) + " and " +
                       "new_equipment_info.Equip_status = 'Ok' and new_equipment_info.Unit_name = 'Beas Hosp'and new_equipment_info.Category= '" + Cat + "' ORDER BY t1.RDATE desc"; 
                }
            
                con.Open();              

                SqlCommand cmd = new SqlCommand(p, con);                
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "Repair_master");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "Repair_master";
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Overdueload();

            }
        }

        private void btnCons_Click(object sender, EventArgs e)
        {
            ButtonClick();
            
            btnCons.ForeColor = System.Drawing.Color.Red;
          
            Consumption = true;            
            groupBoxConsumption.Visible = true;
           

            using (SqlConnection con = new SqlConnection(connString))
            {

                string p = ("Select * from consumption  order by DOC_DT desc");
                SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                DataSet ds = new DataSet();
                dataadapter.Fill(ds, "consumption");
                EquipGrid.DataSource = ds;
                EquipGrid.DataMember = "consumption";
                //count
                string qty = Convert.ToString(EquipGrid.Rows.Count);
                lblCount.Text = qty;
                Consumpload();
            }
        }

        private void btnSearchRepair_Click(object sender, EventArgs e)
        {
            

                using (SqlConnection con = new SqlConnection(connString))
                {
                    //SELECT Orders.OrderID, Customers.CustomerName FROM Orders iNNER JOIN Customers ON Orders.CustomerID = Customers.CustomerID;
                    string p = ("Select  Repair_Master.RID,Repair_Master.Rdate,Repair_Master.Eq_ID,new_equipment_info.Equipment_Name,new_equipment_info.Complete_specification,Repair_Master.Dept,Repair_Master.Repair_maint,Repair_Master.Action_taken,Repair_Master.Sapres,Repair_Master.Status,Repair_Master.Att_by from Repair_Master iNNER JOIN New_Equipment_info ON Repair_Master.Eq_id = New_Equipment_info.Equipment_Id  " +
                        "where Equipment_Name LIKE '%" + txtSearchRepair.Text + "%' order by Rdate desc");
                    SqlDataAdapter dataadapter = new SqlDataAdapter(p, con);
                    DataTable Dt = new DataTable();
                    dataadapter.Fill(Dt);
                    //EquipGrid.Refresh();
                    //EquipGrid.Rows.Clear();              

                    EquipGrid.DataSource = Dt;
                    //count
                    string qty = Convert.ToString(EquipGrid.Rows.Count);
                    lblCount.Text = qty;
                    Repairload();


                }

            
        }

        private void Repairload()
        {

            EquipGrid.Columns[0].Width = 50;
            EquipGrid.Columns[1].Width = 90;
            EquipGrid.Columns[2].Width = 120;
            EquipGrid.Columns[3].Width = 150;
            EquipGrid.Columns[4].Width = 400;
            EquipGrid.Columns[5].Width = 100;
            EquipGrid.Columns[6].Width = 100;
            EquipGrid.Columns[7].Width = 300;
            EquipGrid.Columns[8].Width = 100;
            EquipGrid.Columns[9].Width = 100;
            EquipGrid.Columns[10].Width = 100;


        }
        private void Overdueload()
        {

            EquipGrid.Columns[0].Width = 50;
            EquipGrid.Columns[1].Width = 90;
            EquipGrid.Columns[2].Width = 80;
            EquipGrid.Columns[3].Width = 100;
            EquipGrid.Columns[4].Width = 400;
            EquipGrid.Columns[5].Width = 100;
            EquipGrid.Columns[6].Width = 70;
            EquipGrid.Columns[7].Width = 100;
            EquipGrid.Columns[8].Width = 80;
            //EquipGrid.Columns[9].Width = 100;
            //EquipGrid.Columns[10].Width = 100;


        }
        private void Consumpload()
        {
            EquipGrid.Columns[0].Width = 60;
            EquipGrid.Columns[1].Width = 60; 
            EquipGrid.Columns[2].Width = 80;
            EquipGrid.Columns[3].Width = 40;
            EquipGrid.Columns[4].Width = 70;
            EquipGrid.Columns[5].Width = 300;
            EquipGrid.Columns[6].Width = 50;
            EquipGrid.Columns[7].Width = 70;
            EquipGrid.Columns[8].Width = 70;
            EquipGrid.Columns[9].Width = 50;
            EquipGrid.Columns[10].Width = 70;
            EquipGrid.Columns[11].Width = 50;
            EquipGrid.Columns[12].Width = 220;


        }
        private void Purchload()
        {
          
            EquipGrid.Columns[0].Width = 60;
            EquipGrid.Columns[1].Width = 80;
            EquipGrid.Columns[2].Width = 40;
            EquipGrid.Columns[3].Width = 70;
            EquipGrid.Columns[4].Width = 300;
            EquipGrid.Columns[5].Width = 50;
            EquipGrid.Columns[6].Width = 70;
            EquipGrid.Columns[7].Width = 70;
            EquipGrid.Columns[8].Width = 50;
            EquipGrid.Columns[9].Width = 70;
            EquipGrid.Columns[10].Width = 50;
            EquipGrid.Columns[11].Width = 220;
        }
            private void Maintload()
        {
            EquipGrid.Columns[0].Width = 50;
            EquipGrid.Columns[1].Width = 90;
            EquipGrid.Columns[2].Width = 120;
            EquipGrid.Columns[3].Width = 400;
            EquipGrid.Columns[4].Width = 100;
            EquipGrid.Columns[5].Width = 100;
            EquipGrid.Columns[6].Width = 80;
            EquipGrid.Columns[7].Width  =80;
            


        }

        private void Stockload()
        {
            EquipGrid.Columns[0].Width = 100;
            EquipGrid.Columns[1].Width = 420;
            EquipGrid.Columns[2].Width = 90;
            EquipGrid.Columns[3].Width = 90;
            //EquipGrid.Columns[4].Width = 100;
            //EquipGrid.Columns[5].Width = 200;
            //EquipGrid.Columns[6].Width = 100;
            //EquipGrid.Columns[7].Width = 150;


        }
        private void Equipload()
        {            
            EquipGrid.Columns[0].Width = 40;
            EquipGrid.Columns[1].Width = 100;
            EquipGrid.Columns[2].Width = 120;
            EquipGrid.Columns[3].Width = 300;
            EquipGrid.Columns[4].Width = 100;
            EquipGrid.Columns[5].Width = 80;
            EquipGrid.Columns[6].Width = 80;
            EquipGrid.Columns[7].Width = 120;
            EquipGrid.Columns[8].Width = 50;
            EquipGrid.Columns[9].Width = 80;
            
        }
        private void GroupBox()
        {

            groupBoxEquip.Visible = true;
            groupBoxPurc.Visible = false;
            groupBoxConsumption.Visible = false;
            groupBoxRepair.Visible = false;
            groupBoxMaint.Visible = false;
        }

        private void ButtonClick()
        {
            Purchase = Pending = Maintenance = Equipment = Consumption = Repair = EquipName = DeptName = Stock = ReminderTasks = false;
            btnEquip.Text = "    Equipments";
            btnEquip.ForeColor = System.Drawing.Color.Maroon;

            btnCons.Text = "  Consumption";
            btnCons.ForeColor = System.Drawing.Color.Maroon;
            btnCons.TextAlign = ContentAlignment.MiddleRight;

            btnRepair.Text = "Repair";
            btnRepair.ForeColor = System.Drawing.Color.Maroon;

            btnMaint.Text = "Maintenance";
            btnMaint.ForeColor = System.Drawing.Color.Maroon;
            btnMaint.TextAlign = ContentAlignment.MiddleRight;

            btnPurch.Text = "  Purchase";
            btnPurch.ForeColor = System.Drawing.Color.Maroon;

            btnStock.Text = "Stock";
            btnStock.ForeColor = System.Drawing.Color.Maroon;

            btnEquiName.Text = "Equip.Names";
            btnEquiName.ForeColor = System.Drawing.Color.Maroon;
            btnEquiName.TextAlign = ContentAlignment.MiddleRight;


            btnDepName.Text = "Location";
            btnDepName.ForeColor = System.Drawing.Color.Maroon;


            btnPend.Text = "  Pending";
            btnPend.ForeColor = System.Drawing.Color.Maroon;

            btnReminder.Text = "  Reminder";
            btnReminder.ForeColor = System.Drawing.Color.Maroon;

            this.EquipGrid.RowsDefaultCellStyle.BackColor = Color.AliceBlue;
            this.EquipGrid.AlternatingRowsDefaultCellStyle.BackColor =
            Color.Beige;
            // this.EquipGrid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //this.WindowState = FormWindowState.Maximized;


            progressBar.Visible = false;
            lblstatus.Visible = false;
            groupBoxEquip.Visible = false;
            groupBoxPurc.Visible = false;
            groupBoxConsumption.Visible = false;
            groupBoxRepair.Visible = false;
            groupBoxMaint.Visible = false;
            groupBoxStock.Visible = false;
            //groupBox3.Visible = true;
            // position
            EquipGrid.Location = new Point(180, 100); 
            EquipGrid.Size = new Size(1100, 520);

            //position of group box
            groupBoxMaint.Location = new Point(230, 40);
            groupBoxEquip.Location = new Point(230, 40);
            groupBoxRepair.Location = new Point(230, 40);
            groupBoxStock.Location = new Point(230, 40);
            groupBoxConsumption.Location = new Point(230, 40);
            groupBoxPurc.Location = new Point(230, 40);





        }
    }
    }
    







