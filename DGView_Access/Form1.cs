using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using MetroFramework.Forms;

namespace DGView_Access
{
    public partial class Form1 : MetroForm
    {
        private const string conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:/C#/GridView_CRUD/spacecraftsDB.mdb;";
        readonly OleDbConnection con = new OleDbConnection(conString);
        OleDbCommand cmd;
        OleDbDataAdapter adapter;
        readonly DataTable dt = new DataTable();

      /*
     * CONSTRUCTOR
     */
        public Form1()
        {
            InitializeComponent();
            //DATAGRIDVIEW PROPERTIES
            dataGridView1.ColumnCount = 5;
            dataGridView1.Columns[0].Name = "ID";
            dataGridView1.Columns[1].Name = "*";
            dataGridView1.Columns[2].Name = "Name";
            dataGridView1.Columns[3].Name = "Propellant";
            dataGridView1.Columns[4].Name = "Destination";

            // SETTING HEADER WIDTH
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //SELECTION MODE
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;

            // Call load data
            retrieve();
        }

        /*
         * INSERT INTO DB
         */
        private void add(string name, string propellant, string destination)
        {
            //SQL STMT
            const string sql = "INSERT INTO spacecraftsTB(S_Name,S_Propellant,S_Destination) VALUES(@NAME,@PROPELLANT,@DESTINATION)";
            cmd = new OleDbCommand(sql, con);

            //ADD PARAMS
            cmd.Parameters.AddWithValue("@NAME", name);
            cmd.Parameters.AddWithValue("@PROPELLANT", propellant);
            cmd.Parameters.AddWithValue("@DESTINATION", destination);

            //OPEN CON AND EXEC INSERT
            try
            {
                con.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                }
                con.Close();
                retrieve();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        /*
         * FILL DATAGRIDVIEW
         */
        private void populate(string id, string name, string propellant, string destination)
        {
            dataGridView1.Rows.Add(id, null,name, propellant, destination);
        }

        /*
         * RETRIEVAL OF DATA
         */
        private void retrieve()
        {
            dataGridView1.Rows.Clear();
            //SQL STATEMENT
            String sql = "SELECT * FROM spacecraftsTB ";
            cmd = new OleDbCommand(sql, con);
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
                //LOOP THROUGH DATATABLE
                foreach (DataRow row in dt.Rows)
                {
                    populate(row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString());
                }

                con.Close();
                //CLEAR DATATABLE 
                dt.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        /*
          *  UPDATE DATABASE
         */
        private void update(int id, string name, string propellant, string destination)
        {
            //SQL STATEMENT
            string sql = "UPDATE spacecraftsTB SET S_Name='" + name + "',S_Propellant='" + propellant + "',S_Destination='" + destination + "' WHERE ID=" + id + "";
            cmd = new OleDbCommand(sql, con);

            //OPEN CONNECTION,UPDATE,RETRIEVE DATAGRIDVIEW
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd)
                {
                    UpdateCommand = con.CreateCommand()
                };
                adapter.UpdateCommand.CommandText = sql;
                if (adapter.UpdateCommand.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                }
                con.Close();

                //REFRESH DATA
                retrieve();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        /*
         * DELETE FROM DATABASE
         */
        private void delete(int id)
        {
            String sql = "";
            //SQL STATEMENT CREATE
            if (id == -1)
            {
                sql = "DELETE FROM spacecraftsTB";
            }
            else
            {
                sql = "DELETE FROM spacecraftsTB WHERE ID=" + id + "";
            }
            cmd = new OleDbCommand(sql, con);

            //'OPEN CONNECTION,EXECUTE DELETE,CLOSE CONNECTION
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.DeleteCommand = con.CreateCommand();
                adapter.DeleteCommand.CommandText = sql;

                //PROMPT FOR CONFIRMATION BEFORE DELETING
                if (MessageBox.Show(@"Are you sure to permanently delete this?", @"DELETE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show(@"Successfully deleted");
                    }
                }
                con.Close();
                retrieve();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }
        /*
         * CLEAR TEXTBOXES
         */
        private void clearTxts()
        {
            nameTxt.Text = "";
            propellantTxt.Text = "";
            destinationTxt.Text = "";
        }

        /*
         * CHECK EDITING DATA IS EXISTING
         */
        private bool checkEditing()
        {
            // AFTER CLEARING DATA CHECK
            if (dataGridView1.Rows.Count <= 1 && dataGridView1.Rows[0].Cells[0].Value == null)
            {
                return false;
            }

            // FOCUSING INDEX IS AN EXISTING ROW
            if (dataGridView1.SelectedRows[0].Cells[0].Value == null)
            {
                return false;
            }
            // GET DATA ON INPUT AREA
            String sName = nameTxt.Text;
            String spropellant = propellantTxt.Text;
            String sdestination = destinationTxt.Text;
            // GET DATA ON FOCUSING ROW
            String sGridName = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            String sGridpropellant = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
            String sGriddestination = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            if (sName != sGridName || spropellant != sGridpropellant || sdestination != sGriddestination ) {
                return true;
            }
            return false;
        }

        /*
         * ADD BUTTON CLICKED
         */
        private void addBtn_Click(object sender, EventArgs e)
        {
            // CHECK DATA EMPTY
            if (nameTxt.Text != null || propellantTxt.Text != null || destinationTxt.Text != null) {
                add(nameTxt.Text, propellantTxt.Text, destinationTxt.Text);
            }
        }

        /*
         * RETRIEVE BUTTON CLICKED
         */
        private void retrieveBtn_Click(object sender, EventArgs e)
        {
            if (checkEditing()) {
                //PROMPT FOR CONFIRMATION BEFORE RETRIEVE
                if (MessageBox.Show(@"Having edited data, Are you sure to reload?", @"RELOAD", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    retrieve();
                }
            }
        }

        /*
         * UPDATE BUTTON CLICKED
         */
        private void updateBtn_Click(object sender, EventArgs e)
        {
            if (checkEditing())
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;
                if (selectedIndex != -1)
                {
                    if (dataGridView1.SelectedRows[0].Cells[0].Value == null)
                    {
                        // Do nothing
                    }
                    else
                    {
                        String selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                        int id = Convert.ToInt32(selected);
                        update(id, nameTxt.Text, propellantTxt.Text, destinationTxt.Text);
                    }
                }
            }
        }
        /*
         * DELETE BUTTON CLICKED
         */
        private void deleteBtn_Click(object sender, EventArgs e)
        {
            int selectedIndex = dataGridView1.SelectedRows[0].Index;
            if (selectedIndex != -1 && selectedIndex + 1 < dataGridView1.Rows.Count)
            {
                // Delete creating new row
                if (dataGridView1.SelectedRows[0].Cells[0].Value == null) {
                    dataGridView1.Rows.RemoveAt(selectedIndex);
                } else {
                    // Delete saved row
                    int id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[0].Value.ToString());
                    delete(id);
                }
            }
        }
        /*
         * CLEAR BUTTON CLICKED
         */
        private void clearBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(@"Did you want to clear all data?", @"RELOAD", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                dataGridView1.Rows.Clear();
                clearTxts();
                delete(-1);
            }
        }
       
       

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int selectedIndex = dataGridView1.SelectedRows[0].Index;
                if (selectedIndex != -1)
                {
                    if (dataGridView1.SelectedRows[0].Cells[0].Value != null)
                    {
                        string name = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                        string propellant = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                        string destination = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                        // REMOVE EVENT HANDLE
                        nameTxt.TextChanged -= input_TextChanged;
                        propellantTxt.TextChanged -= input_TextChanged;
                        destinationTxt.TextChanged -= input_TextChanged;
                        // SET DATA
                        nameTxt.Text = name;
                        propellantTxt.Text = propellant;
                        destinationTxt.Text = destination;
                        // ADD EVENT HANDLE
                        nameTxt.TextChanged += input_TextChanged;
                        propellantTxt.TextChanged += input_TextChanged;
                        destinationTxt.TextChanged += input_TextChanged;
                    }

                } else
                {
                    updateBtn.Enabled = false;
                    deleteBtn.Enabled = false;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
            }
        }

        private void input_TextChanged(object sender, EventArgs e)
        {
            // CHECK EDITING DATA
            if (checkEditing())
            {
                // ABLE EDITING MARK
                dataGridView1.SelectedRows[0].Cells[1].Value = "*";
            }
        }
    }
}