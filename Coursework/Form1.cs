using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Aspose.Pdf;
using System.IO;

namespace Coursework
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ConnectDB();
            tabControl1.Height = this.Height - 110;
        }
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GameDB.accdb;Persist Security Info=False;";
        public OleDbConnection connection;
        public string query;
        public OleDbCommand command;
        public OleDbDataAdapter adapter;
        public DataTable dataTable;
        OleDbDataAdapter table1, table2, table3, table4, table5, table6, table7, table8, table9, table10, table11, table12, table13, table14;

        DataTable table1DS, table2DS, table3DS, table4DS, table5DS, table6DS, table7DS, table8DS, table9DS, table10DS, table11DS, table12DS, table13DS, table14DS;

        public void ConnectDB()
        {
            connection = new OleDbConnection(connectString);
            connection.Open();

            string sqlTable1 = "SELECT * FROM account";
            string sqlTable2 = "SELECT * FROM account_roles";
            string sqlTable3 = "SELECT * FROM role";
            string sqlTable4 = "SELECT * FROM account_characters";
            string sqlTable5 = "SELECT * FROM characters";
            string sqlTable6 = "SELECT * FROM character_achievements";
            string sqlTable7 = "SELECT * FROM achievement";
            string sqlTable8 = "SELECT * FROM transactions";
            string sqlTable9 = "SELECT * FROM store_item";
            string sqlTable10 = "SELECT * FROM ticket_authors";
            string sqlTable11 = "SELECT * FROM ticket";
            string sqlTable12 = "SELECT * FROM ticket_messages";
            string sqlTable13 = "SELECT * FROM message";
            string sqlTable14 = "SELECT * FROM message_authors";
            table1 = new OleDbDataAdapter(sqlTable1, connection);
            table2 = new OleDbDataAdapter(sqlTable2, connection);
            table3 = new OleDbDataAdapter(sqlTable3, connection);
            table4 = new OleDbDataAdapter(sqlTable4, connection);
            table5 = new OleDbDataAdapter(sqlTable5, connection);
            table6 = new OleDbDataAdapter(sqlTable6, connection);
            table7 = new OleDbDataAdapter(sqlTable7, connection);
            table8 = new OleDbDataAdapter(sqlTable8, connection);
            table9 = new OleDbDataAdapter(sqlTable9, connection);
            table10 = new OleDbDataAdapter(sqlTable10, connection);
            table11 = new OleDbDataAdapter(sqlTable11, connection);
            table12 = new OleDbDataAdapter(sqlTable12, connection);
            table13 = new OleDbDataAdapter(sqlTable13, connection);
            table14 = new OleDbDataAdapter(sqlTable14, connection);
            table1DS = new DataTable();
            table2DS = new DataTable();
            table3DS = new DataTable();
            table4DS = new DataTable();
            table5DS = new DataTable();
            table6DS = new DataTable();
            table7DS = new DataTable();
            table8DS = new DataTable();
            table9DS = new DataTable();
            table10DS = new DataTable();
            table11DS = new DataTable();
            table12DS = new DataTable();
            table13DS = new DataTable();
            table14DS = new DataTable();
            table1.Fill(table1DS);
            table2.Fill(table2DS);
            table3.Fill(table3DS);
            table4.Fill(table4DS);
            table5.Fill(table5DS);
            table6.Fill(table6DS);
            table7.Fill(table7DS);
            table8.Fill(table8DS);
            table9.Fill(table9DS);
            table10.Fill(table10DS);
            table11.Fill(table11DS);
            table12.Fill(table12DS);
            table13.Fill(table13DS);
            table14.Fill(table14DS);
            table2DS.Columns[0].AutoIncrement = true;
            table3DS.Columns[0].AutoIncrement = true;
            table4DS.Columns[0].AutoIncrement = true;
            table5DS.Columns[0].AutoIncrement = true;
            table6DS.Columns[0].AutoIncrement = true;
            table7DS.Columns[0].AutoIncrement = true;
            table8DS.Columns[0].AutoIncrement = true;
            table9DS.Columns[0].AutoIncrement = true;
            table10DS.Columns[0].AutoIncrement = true;
            table11DS.Columns[0].AutoIncrement = true;
            table12DS.Columns[0].AutoIncrement = true;
            table13DS.Columns[0].AutoIncrement = true;
            table14DS.Columns[0].AutoIncrement = true;

            dataGridView1.DataSource = table1DS;
            dataGridView2.DataSource = table2DS;
            dataGridView3.DataSource = table3DS;
            dataGridView4.DataSource = table4DS;
            dataGridView5.DataSource = table5DS;
            dataGridView6.DataSource = table6DS;
            dataGridView7.DataSource = table7DS;
            dataGridView8.DataSource = table8DS;
            dataGridView9.DataSource = table9DS;
            dataGridView10.DataSource = table10DS;
            dataGridView11.DataSource = table11DS;
            dataGridView12.DataSource = table12DS;
            dataGridView13.DataSource = table13DS;
            dataGridView14.DataSource = table14DS;


        }
        public void CloseDB()
        {
            connection.Close();
        }
        private void SetTables()
        {
            string sqlTable1 = "SELECT * FROM account";
            string sqlTable2 = "SELECT * FROM account_roles";
            string sqlTable3 = "SELECT * FROM role";
            string sqlTable4 = "SELECT * FROM account_characters";
            string sqlTable5 = "SELECT * FROM characters";
            string sqlTable6 = "SELECT * FROM character_achievements";
            string sqlTable7 = "SELECT * FROM achievement";
            string sqlTable8 = "SELECT * FROM transactions";
            string sqlTable9 = "SELECT * FROM store_item";
            string sqlTable10 = "SELECT * FROM ticket_authors";
            string sqlTable11 = "SELECT * FROM ticket";
            string sqlTable12 = "SELECT * FROM ticket_messages";
            string sqlTable13 = "SELECT * FROM message";
            string sqlTable14 = "SELECT * FROM message_authors";
            table1 = new OleDbDataAdapter(sqlTable1, connection);
            table2 = new OleDbDataAdapter(sqlTable2, connection);
            table3 = new OleDbDataAdapter(sqlTable3, connection);
            table4 = new OleDbDataAdapter(sqlTable4, connection);
            table5 = new OleDbDataAdapter(sqlTable5, connection);
            table6 = new OleDbDataAdapter(sqlTable6, connection);
            table7 = new OleDbDataAdapter(sqlTable7, connection);
            table8 = new OleDbDataAdapter(sqlTable8, connection);
            table9 = new OleDbDataAdapter(sqlTable9, connection);
            table10 = new OleDbDataAdapter(sqlTable10, connection);
            table11 = new OleDbDataAdapter(sqlTable11, connection);
            table12 = new OleDbDataAdapter(sqlTable12, connection);
            table13 = new OleDbDataAdapter(sqlTable13, connection);
            table14 = new OleDbDataAdapter(sqlTable14, connection);
        }

        public void RefreshTables()
        {
            table1DS.Rows.Clear();
            table1.Fill(table1DS);
            table2DS.Rows.Clear();
            table2.Fill(table2DS);
            table3DS.Rows.Clear();
            table3.Fill(table3DS);
            table4DS.Rows.Clear();
            table4.Fill(table4DS);
            table5DS.Rows.Clear();
            table5.Fill(table5DS);
            table6DS.Rows.Clear();
            table6.Fill(table6DS);
            table7DS.Rows.Clear();
            table7.Fill(table7DS);
            table8DS.Rows.Clear();
            table8.Fill(table8DS);
            table9DS.Rows.Clear();
            table9.Fill(table9DS);
            table10DS.Rows.Clear();
            table10.Fill(table10DS);
            table11DS.Rows.Clear();
            table11.Fill(table11DS);
            table12DS.Rows.Clear();
            table12.Fill(table12DS);
            table13DS.Rows.Clear();
            table13.Fill(table13DS);
            table14DS.Rows.Clear();
            table14.Fill(table14DS);


        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            tabControl1.Height = this.Height - 110;
        }
        private void tabControl3_Selected(object sender, TabControlEventArgs e)
        {
            int index = tabControl3.SelectedIndex;
            switch (index)
            {
                case 1:
                    string input = Interaction.InputBox("Input Login", "Login", "login1234");
                    query = "SELECT account.user_account_login, account.user_account_name, characters.character_name, characters.character_class, characters.character_race FROM characters INNER JOIN (account INNER JOIN account_characters ON account.user_account_login = account_characters.[user_account_login]) ON characters.character_id = account_characters.character_id WHERE account.user_account_login='" + input + "' ORDER BY characters.character_name ASC;";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    dataTable.Rows.Clear();
                    adapter.Fill(dataTable);
                    dataGridView15.DataSource = null;
                    dataGridView15.DataSource = dataTable;
                    break;
                case 2:
                    //Prompt Form
                    Form promptQuery = new Form();
                    promptQuery.FormBorderStyle = FormBorderStyle.FixedDialog;
                    promptQuery.StartPosition = FormStartPosition.CenterScreen;
                    promptQuery.MaximizeBox = false;
                    promptQuery.MinimizeBox = false;
                    promptQuery.Width = 500;
                    promptQuery.Height = 300;
                    promptQuery.Text = "Query";
                    GroupBox Radios2 = new GroupBox();
                    Radios2.Width = 500;
                    Radios2.Height = 300;
                    Radios2.Location = new System.Drawing.Point(0, 0);
                    Radios2.BackColor = System.Drawing.Color.Transparent;
                    RadioButton BiggerButton1 = new RadioButton() { Text = "Older", Left = promptQuery.Width / 2 - 175, Top = 25, Name = "QueryAccountChoice", BackColor = System.Drawing.Color.Transparent };
                    RadioButton SmallerButton1 = new RadioButton() { Text = "Younger", Left = promptQuery.Width / 2 - 50, Top = 25, Name = "QueryAccountChoice", BackColor = System.Drawing.Color.Transparent };
                    RadioButton EqualsButton1 = new RadioButton() { Text = "Equals", Left = promptQuery.Width / 2 + 75, Top = 25, Name = "QueryAccountChoice", BackColor = System.Drawing.Color.Transparent };
                    RadioButton BiggerButton2 = new RadioButton() { Text = "Older", Left = promptQuery.Width / 2 - 175, Top = 100, Name = "QueryCharacterChoice", BackColor = System.Drawing.Color.Transparent };
                    RadioButton SmallerButton2 = new RadioButton() { Text = "Younger", Left = promptQuery.Width / 2 - 50, Top = 100, Name = "QueryCharacterChoice", BackColor = System.Drawing.Color.Transparent };
                    RadioButton EqualsButton2 = new RadioButton() { Text = "Equals", Left = promptQuery.Width / 2 + 75, Top = 100, Name = "QueryCharacterChoice", BackColor = System.Drawing.Color.Transparent };
                    Button confirm = new Button() { Text = "Confirm", Left = promptQuery.Width / 2 - 150, Width = 100, Top = 200, Height = 50, DialogResult = DialogResult.OK };
                    Button cancel = new Button() { Text = "Cancel", Left = promptQuery.Width / 2 + 50, Width = 100, Top = 200, Height = 50, DialogResult = DialogResult.Cancel };
                    Label promptLabelCharacterAge = new Label() { Text = "Input Character Date", Left = promptQuery.Width / 2 - 150, Top = 125, BackColor = System.Drawing.Color.Transparent };
                    Label promptLabelAccountAge = new Label() { Text = "Input Account Date", Left = promptQuery.Width / 2 - 150, Top = 50, BackColor = System.Drawing.Color.Transparent };
                    DateTimePicker characterDate = new DateTimePicker() { Left = promptQuery.Width / 2 - 150, Top = 150, Width = promptQuery.Width / 2 };
                    DateTimePicker accountDate = new DateTimePicker() { Left = promptQuery.Width / 2 - 150, Top = 75, Width = promptQuery.Width / 2 };

                    promptQuery.Controls.Add(confirm);
                    promptQuery.Controls.Add(cancel);
                    promptQuery.Controls.Add(promptLabelCharacterAge);
                    promptQuery.Controls.Add(promptLabelAccountAge);
                    promptQuery.Controls.Add(characterDate);
                    promptQuery.Controls.Add(accountDate);

                    promptQuery.Controls.Add(BiggerButton1);
                    promptQuery.Controls.Add(SmallerButton1);
                    promptQuery.Controls.Add(EqualsButton1);

                    Radios2.Controls.Add(BiggerButton2);
                    Radios2.Controls.Add(SmallerButton2);
                    Radios2.Controls.Add(EqualsButton2);
                    promptQuery.Controls.Add(Radios2);
                    promptQuery.ShowDialog();
                    if (promptQuery.DialogResult != DialogResult.OK)
                    {
                        promptQuery.Close();
                        return;
                    }
                    string characterDateVal = characterDate.Value.ToShortDateString();
                    string accountDateVal = accountDate.Value.ToShortDateString();
                    string characterSign = ">", accountSign = ">";
                    if (SmallerButton1.Checked)
                    {
                        accountSign = ">";

                    }
                    if (BiggerButton1.Checked)
                    {
                        accountSign = "<";

                    }
                    if (EqualsButton1.Checked)
                    {
                        accountSign = "=";
                    }
                    if (SmallerButton2.Checked)
                    {
                        characterSign = ">";

                    }
                    if (BiggerButton2.Checked)
                    {
                        characterSign = "<";

                    }
                    if (EqualsButton2.Checked)
                    {
                        characterSign = "=";
                    }
                    query = "SELECT account.user_account_login, account.user_account_dor, Count(characters.character_id) AS [Number of Characters] FROM characters INNER JOIN (account INNER JOIN account_characters ON account.user_account_login = account_characters.user_account_login) ON characters.character_id = account_characters.character_id WHERE account.user_account_dor" + accountSign + " #" + accountDateVal + "# AND account_characters.player_character_doc" + characterSign + " #" + characterDateVal + "# GROUP BY account.user_account_login, account.user_account_dor;";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    dataTable.Rows.Clear();
                    adapter.Fill(dataTable);
                    dataGridView16.DataSource = dataTable;
                    break;
                case 3:
                    query = "SELECT account.user_account_login, Sum(store_item.store_item_price) AS [Account Net Worth] FROM store_item INNER JOIN(account INNER JOIN transactions ON account.user_account_login = transactions.user_account_login) ON store_item.[store_item_id]= Transactions.[store_item_id] GROUP BY account.user_account_login;";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    dataTable.Rows.Clear();
                    adapter.Fill(dataTable);
                    dataGridView17.DataSource = dataTable;
                    break;
                case 4:
                    query = "SELECT account.user_account_login, SUM(store_item.store_item_price) AS [Account Net Worth] FROM store_item INNER JOIN(account INNER JOIN transactions ON account.user_account_login = transactions.user_account_login) ON store_item.store_item_id = transactions.store_item_id GROUP BY account.user_account_login HAVING SUM(store_item.store_item_price) > (SELECT AVG(total_net_worth) FROM(SELECT SUM(store_item.store_item_price) AS total_net_worth FROM store_item INNER JOIN transactions ON store_item.store_item_id = transactions.store_item_id GROUP BY transactions.user_account_login) AS subquery);";
                    command = new OleDbCommand(query, connection);
                    adapter = new OleDbDataAdapter(command);
                    dataTable = new DataTable();
                    dataTable.Rows.Clear();
                    adapter.Fill(dataTable);
                    dataGridView18.DataSource = dataTable;

                    break;


            }

        }


        private void DeleteButton_Click(object sender, EventArgs e)
        {
            //Deletes Data Base entry
            int testint = 0;
            bool testconvert;
            int TabIndex = tabControl1.SelectedIndex;
            string indexName = "";
            string indexValue = "";
            string tableName = tabControl1.SelectedTab.Text;
            try
            {
                switch (TabIndex)
                {
                    case 0:
                        indexValue = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView1.Columns[0].HeaderText;
                        break;
                    case 1:
                        indexValue = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView2.Columns[0].HeaderText;
                        break;
                    case 2:
                        indexValue = dataGridView3.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView3.Columns[0].HeaderText;
                        break;
                    case 3:
                        indexValue = dataGridView4.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView4.Columns[0].HeaderText;
                        break;
                    case 4:
                        indexValue = dataGridView5.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView5.Columns[0].HeaderText;
                        break;
                    case 5:
                        indexValue = dataGridView6.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView6.Columns[0].HeaderText;
                        break;
                    case 6:
                        indexValue = dataGridView7.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView7.Columns[0].HeaderText;
                        break;
                    case 7:
                        indexValue = dataGridView8.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView8.Columns[0].HeaderText;
                        break;
                    case 8:
                        indexValue = dataGridView9.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView9.Columns[0].HeaderText;
                        break;
                    case 9:
                        indexValue = dataGridView10.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView10.Columns[0].HeaderText;
                        break;
                    case 10:
                        indexValue = dataGridView11.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView11.Columns[0].HeaderText;
                        break;
                    case 11:
                        indexValue = dataGridView12.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView12.Columns[0].HeaderText;
                        break;
                    case 12:
                        indexValue = dataGridView13.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView13.Columns[0].HeaderText;
                        break;
                    case 13:
                        indexValue = dataGridView14.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView14.Columns[0].HeaderText;
                        break;


                }

                testconvert = int.TryParse(indexValue, out testint);
                if (testconvert == true)
                {
                    query = "DELETE FROM " + tableName + " WHERE [" + indexName + "]=" + indexValue + ";";
                }
                else
                {
                    query = "DELETE FROM " + tableName + " WHERE [" + indexName + "]='" + indexValue + "';";
                }

                OleDbCommand command = new OleDbCommand(query, connection);
                command.ExecuteNonQuery();
                RefreshTables();
            }
            catch
            {
                MessageBox.Show("Error: You are trying to delete while you didn't choose a row!");
            }

        }


        private void EditButton_Click(object sender, EventArgs e)
        {
            //Edits Data Base entry
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            string indexValue = "";
            string indexName = "";
            bool test = false;
            int j = 0;
            List<Control> inputs = new List<Control>();
            List<Label> labels = new List<Label>();
            List<String> columns = new List<String>();
            //Building Prompt
            Form prompt = new Form();
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            prompt.MaximizeBox = false;
            prompt.MinimizeBox = false;
            prompt.Width = 500;
            prompt.Height = 500;
            Button confirm = new Button() { Text = "Confirm", Left = prompt.Width / 2 - 150, Width = 100, Top = prompt.Height - 130, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = prompt.Width / 2 + 50, Width = 100, Top = prompt.Height - 130, Height = 50, DialogResult = DialogResult.Cancel };
            prompt.Controls.Add(confirm);
            prompt.Controls.Add(cancel);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;
            //Checking which table

            try
            {

                switch (tabIndex)

                {
                    case 0:
                        if (!table1DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView1.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.Columns[i].HeaderText.ToString() });
                            if (dataGridView1.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView1.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView1.Columns[0].HeaderText.ToString();
                        break;
                    case 1:
                        if (!table2DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView2.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.Columns[i].HeaderText.ToString() });
                            if (dataGridView2.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView2.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView2.Columns[0].HeaderText.ToString();
                        break;
                    case 2:
                        if (!table3DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView3.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.Columns[i].HeaderText.ToString() });
                            if (dataGridView3.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView3.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView3.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView3.Columns[0].HeaderText.ToString();
                        break;
                    case 3:
                        if (!table4DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView4.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.Columns[i].HeaderText.ToString() });
                            if (dataGridView4.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView4.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView4.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView4.Columns[0].HeaderText.ToString();
                        break;
                    case 4:
                        if (!table5DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView5.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.Columns[i].HeaderText.ToString() });
                            if (dataGridView5.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView5.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView5.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView5.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView5.Columns[0].HeaderText.ToString();
                        break;
                    case 5:
                        if (!table6DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView6.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.Columns[i].HeaderText.ToString() });
                            if (dataGridView6.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView6.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView6.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView6.Columns[0].HeaderText.ToString();
                        break;
                    case 6:
                        if (!table7DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView7.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView7.Columns[i].HeaderText.ToString() });
                            if (dataGridView7.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView7.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView7.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView7.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView7.Columns[0].HeaderText.ToString();
                        break;
                    case 7:
                        if (!table8DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView8.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.Columns[i].HeaderText.ToString() });
                            if (dataGridView8.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView8.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                        }
                        indexValue = dataGridView8.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView8.Columns[0].HeaderText.ToString();
                        break;
                    case 8:
                        if (!table9DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1-j; i < dataGridView9.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.Columns[i].HeaderText.ToString() });
                            if (dataGridView9.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView9.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView9.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView9.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView9.Columns[0].HeaderText.ToString();
                        break;
                    case 9:
                        if (!table10DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView10.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView10.Columns[i].HeaderText.ToString() });
                            if (dataGridView10.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView10.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView10.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView10.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView10.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView10.Columns[0].HeaderText.ToString();
                        break;
                    case 10:
                        if (!table11DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView11.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView11.Columns[i].HeaderText.ToString() });
                            if (dataGridView11.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView11.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView11.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView11.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView11.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView11.Columns[0].HeaderText.ToString();
                        break;
                    case 11:
                        if (!table12DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView12.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView12.Columns[i].HeaderText.ToString() });
                            if (dataGridView12.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView12.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView12.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView12.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView12.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView12.Columns[0].HeaderText.ToString();
                        break;
                    case 12:
                        if (!table13DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView13.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView13.Columns[i].HeaderText.ToString() });
                            if (dataGridView13.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView13.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView13.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView13.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView13.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView13.Columns[0].HeaderText.ToString();
                        break;
                    case 13:
                        if (!table14DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView14.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView14.Columns[i].HeaderText.ToString() });
                            if (dataGridView14.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Value = (DateTime)dataGridView14.SelectedRows[0].Cells[i].Value });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView14.SelectedRows[0].Cells[i].Value.ToString() });
                            }
                            columns.Add(dataGridView14.Columns[i].HeaderText.ToString());

                        }
                        indexValue = dataGridView14.SelectedRows[0].Cells[0].Value.ToString();
                        indexName = dataGridView14.Columns[0].HeaderText.ToString();
                        break;
                }
                foreach (Control input in inputs)
                {
                    prompt.Controls.Add(input);
                }
                foreach (Label label in labels)
                {
                    prompt.Controls.Add(label);
                }
                prompt.ShowDialog();
                if (prompt.DialogResult == DialogResult.OK)
                {
                    test = int.TryParse(indexValue, out _);
                    query = "UPDATE [" + tabControl1.SelectedTab.Text + "] SET ";
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        if (inputs[i].GetType() == typeof(DateTimePicker))
                        {
                            DateTimePicker temp = (DateTimePicker)inputs[i];
                            query += "[" + columns[i] + "] = #" + temp.Value.ToShortDateString() + "#";
                        }
                        else
                        {
                            query += "[" + columns[i] + "] = '" + inputs[i].Text + "'";
                        }
                        if (i != columns.Count() - 1)
                        {
                            query += ", ";
                        }
                    }
                    if (test == true)
                    {
                        query += " WHERE [" + indexName + "] = " + indexValue + ";";
                    }
                    else
                    {
                        query += " WHERE [" + indexName + "] = '" + indexValue + "';";
                    }


                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.ExecuteNonQuery();
                    RefreshTables();
                    prompt.Close();
                    return;
                }
                if (prompt.DialogResult == DialogResult.Cancel)
                {
                    prompt.Close();

                }


            }
            catch
            {
                MessageBox.Show("Error: You are trying to edit while you didn't choose a row!");

            }

        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            string indexValue = "", indexName = "";
            List<Control> inputs = new List<Control>();
            List<Label> labels = new List<Label>();
            List<String> columns = new List<String>();
            int j = 0;
            //Building Prompt
            Form promptAdd = new Form();
            promptAdd.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptAdd.StartPosition = FormStartPosition.CenterScreen;
            promptAdd.MaximizeBox = false;
            promptAdd.MinimizeBox = false;
            promptAdd.Width = 500;
            promptAdd.Height = 500;
            Button confirm = new Button() { Text = "Confirm", Left = promptAdd.Width / 2 - 150, Width = 100, Top = promptAdd.Height - 130, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = promptAdd.Width / 2 + 50, Width = 100, Top = promptAdd.Height - 130, Height = 50, DialogResult = DialogResult.Cancel };
            promptAdd.Controls.Add(confirm);
            promptAdd.Controls.Add(cancel);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;
            try
            {
                switch (tabIndex)
                {
                    case 0:
                        if (!table1DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView1.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView1.Columns[i].HeaderText.ToString() });
                            if (dataGridView1.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }

                            columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                        }
                        break;
                    case 1:
                        if (table2DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView2.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView2.Columns[i].HeaderText.ToString() });
                            if (dataGridView2.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                        }
                        break;
                    case 2:
                        if (table3DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView3.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView3.Columns[i].HeaderText.ToString() });
                            if (dataGridView3.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                        }

                        break;
                    case 3:
                        if (table4DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView4.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView4.Columns[i].HeaderText.ToString() });
                            if (dataGridView4.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                        }

                        break;
                    case 4:
                        if (table5DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView5.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView5.Columns[i].HeaderText.ToString() });
                            if (dataGridView5.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView5.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 5:
                        if (table6DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView6.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView6.Columns[i].HeaderText.ToString() });
                            if (dataGridView6.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                        }

                        break;
                    case 6:
                        if (table7DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView7.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView7.Columns[i].HeaderText.ToString() });
                            if (dataGridView7.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                        }

                        break;
                    case 7:
                        if (table8DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView8.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView8.Columns[i].HeaderText.ToString() });
                            if (dataGridView8.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                        }

                        break;
                    case 8:
                        if (table9DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView9.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView9.Columns[i].HeaderText.ToString() });
                            if (dataGridView9.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView9.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 9:
                        if (table10DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView10.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView10.Columns[i].HeaderText.ToString() });
                            if (dataGridView10.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView10.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 10:
                        if (table11DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView11.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView11.Columns[i].HeaderText.ToString() });
                            if (dataGridView11.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView11.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 11:
                        if (table12DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView12.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView12.Columns[i].HeaderText.ToString() });
                            if (dataGridView12.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView12.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 12:
                        if (table13DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView13.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView13.Columns[i].HeaderText.ToString() });
                            if (dataGridView13.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView13.Columns[i].HeaderText.ToString());

                        }

                        break;
                    case 13:
                        if (table14DS.Columns[0].AutoIncrement)
                        {
                            j = 1;
                        }
                        for (int i = 1 - j; i < dataGridView14.ColumnCount; i++)
                        {
                            labels.Add(new Label() { Left = 50, Width = 200, Height = 25, Top = 70 + (30 * i), Text = dataGridView14.Columns[i].HeaderText.ToString() });
                            if (dataGridView14.Columns[i].ValueType.Name == "DateTime")
                            {
                                inputs.Add(new DateTimePicker() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });

                            }
                            else
                            {
                                inputs.Add(new TextBox() { Left = 200, Width = 200, Height = 25, Top = 70 + (30 * i) });
                            }
                            columns.Add(dataGridView14.Columns[i].HeaderText.ToString());

                        }

                        break;

                }
                foreach (Control input in inputs)
                {
                    promptAdd.Controls.Add(input);
                }
                foreach (Label label in labels)
                {
                    promptAdd.Controls.Add(label);
                }
                promptAdd.ShowDialog();

                if (promptAdd.DialogResult == DialogResult.OK)
                {
                    query = "INSERT INTO [" + tabControl1.SelectedTab.Text + "] (";
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        query += "[" + columns[i] + "]";

                        if (i != columns.Count() - 1)
                        {
                            query += ", ";
                        }
                    }
                    query += ")  VALUES (";
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        if (inputs[i].GetType() == typeof(DateTimePicker))
                        {
                            DateTimePicker temp = (DateTimePicker)inputs[i];
                            query += "#" + temp.Value.ToShortDateString() + "#";
                        }
                        else
                        {
                            query += "'" + inputs[i].Text + "'";
                        }
                        
                        if (i != columns.Count() - 1)
                        {
                            query += ", ";
                        }
                    }
                    query += ");";

                    OleDbCommand command = new OleDbCommand(query, connection);
                    command.ExecuteNonQuery();
                    RefreshTables();
                    promptAdd.Close();
                    return;
                }
            }
            catch
            {
                MessageBox.Show("Error: Invalid Data Type");
            }

            if (promptAdd.DialogResult == DialogResult.Cancel)
            {
                promptAdd.Close();

            }
        }
        private void searchButton_Click(object sender, EventArgs e)
        {
            //First check if anything is written in the search field
            if (searchBox.Text == "")
            {
                MessageBox.Show("Error, search field cannot be empty!");
                return;
            }

            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            List<String> columns = new List<String>();
            bool isString = true, isInt = false, isDate = false;
            string text = searchBox.Text;
            //New Table for Searching
            OleDbDataAdapter tableSearch;
            DataTable tableSearchData = new DataTable();
            //New Form for showing results of search
            DataGridView searchView = new DataGridView();
            Form SearchResult = new Form();
            SearchResult.FormBorderStyle = FormBorderStyle.FixedDialog;
            SearchResult.StartPosition = FormStartPosition.CenterScreen;
            SearchResult.Width = 750;
            SearchResult.Height = 500;
            searchView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            searchView.Dock = DockStyle.Fill;
            searchView.ReadOnly = true;
            searchView.AllowUserToAddRows = false;
            searchView.AllowUserToDeleteRows = false;
            SearchResult.Controls.Add(searchView);
            if (int.TryParse(text, out _))
            {
                isInt = true;
                isString = false;
                isDate = false;
            }
            else if (System.DateTime.TryParse(text, out _))
            {
                isInt = false;
                isString = false;
                isDate = true;
            }

            switch (tabIndex)

            {
                case 0:

                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView1.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView1.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView1.Columns[i].HeaderText.ToString());
                        }
                    }
                    break;
                case 1:
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView2.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView2.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView2.Columns[i].HeaderText.ToString());
                        }
                    }
                    break;
                case 2:
                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                    {
                        if (dataGridView3.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView3.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView3.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView3.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 3:
                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                    {
                        if (dataGridView4.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView4.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView4.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView4.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 4:
                    for (int i = 0; i < dataGridView5.ColumnCount; i++)
                    {
                        if (dataGridView5.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView5.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView5.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView5.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 5:
                    for (int i = 0; i < dataGridView6.ColumnCount; i++)
                    {
                        if (dataGridView6.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView6.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView6.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView6.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 6:
                    for (int i = 0; i < dataGridView7.ColumnCount; i++)
                    {
                        if (dataGridView7.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView7.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView7.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView7.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 7:
                    for (int i = 0; i < dataGridView8.ColumnCount; i++)
                    {
                        if (dataGridView8.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView8.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView8.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView8.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 8:
                    for (int i = 0; i < dataGridView9.ColumnCount; i++)
                    {
                        if (dataGridView9.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView9.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView9.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView9.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 9:
                    for (int i = 0; i < dataGridView10.ColumnCount; i++)
                    {
                        if (dataGridView10.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView10.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView10.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView10.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 10:
                    for (int i = 0; i < dataGridView11.ColumnCount; i++)
                    {
                        if (dataGridView11.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView11.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView11.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView11.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 11:
                    for (int i = 0; i < dataGridView12.ColumnCount; i++)
                    {
                        if (dataGridView12.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView12.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView12.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView12.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 12:
                    for (int i = 0; i < dataGridView13.ColumnCount; i++)
                    {
                        if (dataGridView13.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView13.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView13.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView13.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
                case 13:
                    for (int i = 0; i < dataGridView14.ColumnCount; i++)
                    {
                        if (dataGridView14.Columns[i].ValueType == Type.GetType("System.String") && isString == true || dataGridView14.Columns[i].ValueType == Type.GetType("System.Int32") && isInt == true || dataGridView14.Columns[i].ValueType == Type.GetType("System.DateTime") && isDate == true)
                        {
                            columns.Add(dataGridView14.Columns[i].HeaderText.ToString());
                        }
                    }

                    break;
            }
            query = "SELECT * FROM [" + tabControl1.SelectedTab.Text.ToString() + "] WHERE ";
            for (int i = 0; i < columns.Count; i++)
            {
                if (isInt == true)
                {
                    query += "[" + columns[i] + "] = " + searchBox.Text + "";
                }
                else
                {
                    query += "[" + columns[i] + "] = '" + searchBox.Text + "'";
                }

                if (i != columns.Count - 1)
                {
                    query += " OR ";
                }
            }
            query += ";";
            searchView.DataSource = null;
            tableSearch = new OleDbDataAdapter(query, connection);
            tableSearch.Fill(tableSearchData);
            searchView.DataSource = tableSearchData;
            SearchResult.Show();
            RefreshTables();
        }
        private void filterButton_Click(object sender, EventArgs e)
        {
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            Type ColumnValueType = null;
            string column = "";
            string FilterValue;
            bool isString = true, isInt = false, isDate = false;
            string filterSign = "";
            //Prompt Form
            Form promptFilter = new Form();
            promptFilter.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptFilter.StartPosition = FormStartPosition.CenterScreen;
            promptFilter.MaximizeBox = false;
            promptFilter.MinimizeBox = false;
            promptFilter.Width = 500;
            promptFilter.Height = 250;
            promptFilter.Text = "Filter";
            RadioButton BiggerButton = new RadioButton()
            {
                Text =
            "Bigger",
                Left = promptFilter.Width / 2 - 175,
                Top = 50,
                Name =
            "FilterChoice",
                BackColor = System.Drawing.Color.Transparent
            };
            RadioButton SmallerButton = new RadioButton()
            {
                Text =
            "Smaller",
                Left = promptFilter.Width / 2 - 50,
                Top = 50,
                Name =
            "FilterChoice",
                BackColor = System.Drawing.Color.Transparent
            };
            RadioButton EqualsButton = new RadioButton()
            {
                Text =
            "Equals",
                Left = promptFilter.Width / 2 + 75,
                Top = 50,
                Name =
            "FilterChoice",
                BackColor = System.Drawing.Color.Transparent
            };
            Button confirm = new Button()
            {
                Text = "Confirm",
                Left =
            promptFilter.Width / 2 - 150,
                Width = 100,
                Top = promptFilter.Height -
            100,
                Height = 50,
                DialogResult = DialogResult.OK
            };
            Button cancel = new Button()
            {
                Text = "Cancel",
                Left =
            promptFilter.Width / 2 + 50,
                Width = 100,
                Top = promptFilter.Height -
            100,
                Height = 50,
                DialogResult = DialogResult.Cancel
            };
            Label promptLabel = new Label()
            {
                Text = "Input Value",
                Left =
            promptFilter.Width / 2 - 150,
                Top = promptFilter.Height - 175,
                BackColor = System.Drawing.Color.Transparent
            };
            TextBox promptText = new TextBox()
            {
                Left =
            promptFilter.Width / 2 - 150,
                Top = promptFilter.Height - 150,
                Width =
            promptFilter.Width / 2
            };
            promptFilter.Controls.Add(confirm);
            promptFilter.Controls.Add(cancel);
            promptFilter.Controls.Add(promptLabel);
            promptFilter.Controls.Add(promptText);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;
            switch (tabIndex)
            {
                case 0:
                    column = dataGridView1.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView1.CurrentCell.OwningColumn.ValueType;
                    break;
                case 1:
                    column = dataGridView2.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView2.CurrentCell.OwningColumn.ValueType;
                    break;
                case 2:
                    column = dataGridView3.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView3.CurrentCell.OwningColumn.ValueType;
                    break;
                case 3:
                    column = dataGridView4.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView4.CurrentCell.OwningColumn.ValueType;
                    break;
                case 4:
                    column = dataGridView5.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView5.CurrentCell.OwningColumn.ValueType;
                    break;
                case 5:
                    column = dataGridView6.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView6.CurrentCell.OwningColumn.ValueType;
                    break;
                case 6:
                    column = dataGridView7.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView7.CurrentCell.OwningColumn.ValueType;
                    break;
                case 7:
                    column = dataGridView8.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView8.CurrentCell.OwningColumn.ValueType;
                    break;
                case 8:
                    column = dataGridView9.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView9.CurrentCell.OwningColumn.ValueType;
                    break;
                case 9:
                    column = dataGridView10.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView10.CurrentCell.OwningColumn.ValueType;
                    break;
                case 10:
                    column = dataGridView11.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView11.CurrentCell.OwningColumn.ValueType;
                    break;
                case 11:
                    column = dataGridView12.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView12.CurrentCell.OwningColumn.ValueType;
                    break;
                case 12:
                    column = dataGridView13.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView13.CurrentCell.OwningColumn.ValueType;
                    break;
                case 13:
                    column = dataGridView14.CurrentCell.OwningColumn.HeaderText;
                    ColumnValueType = dataGridView14.CurrentCell.OwningColumn.ValueType;
                    break;
            }
            if (ColumnValueType != Type.GetType("System.String"))
            {
                promptFilter.Controls.Add(BiggerButton);
                promptFilter.Controls.Add(SmallerButton);
                promptFilter.Controls.Add(EqualsButton);
            }
            promptFilter.ShowDialog();
            if (promptFilter.DialogResult != DialogResult.OK)
            {
                promptFilter.Close();
                return;
            }
            FilterValue = promptText.Text;
            if (ColumnValueType != Type.GetType("System.String"))
            {
                if (BiggerButton.Checked)
                {
                    filterSign = ">";
                }
                if (SmallerButton.Checked)
                {
                    filterSign = "<";
                }
                if (EqualsButton.Checked)
                {
                    filterSign = "=";
                }
                query = "SELECT * FROM [" +
                tabControl1.SelectedTab.Text.ToString() + "] WHERE [" + column + "]" + filterSign + " ";
                if (ColumnValueType == Type.GetType("System.Int32"))
                {
                    query += FilterValue;
                }
                else if (ColumnValueType ==
                Type.GetType("System.DateTime"))
                {
                    query += "#" + FilterValue + "#";
                }
                else
                {
                    query += "'" + FilterValue + "'";
                }
                query += ";";
            }
            else
            {
                query = "SELECT * FROM [" +
                tabControl1.SelectedTab.Text.ToString() + "] WHERE [" + column + "] LIKE '%" + FilterValue + "%'; ";
            }
            switch (tabIndex)
            {
                case 0:
                    table1 = new OleDbDataAdapter(query, connection);
                    break;
                case 1:
                    table2 = new OleDbDataAdapter(query, connection);
                    break;
                case 2:
                    table3 = new OleDbDataAdapter(query, connection);
                    break;
                case 3:
                    table4 = new OleDbDataAdapter(query, connection);
                    break;
                case 4:
                    table5 = new OleDbDataAdapter(query, connection);
                    break;
                case 5:
                    table6 = new OleDbDataAdapter(query, connection);
                    break;
                case 6:
                    table7 = new OleDbDataAdapter(query, connection);
                    break;
                case 7:
                    table8 = new OleDbDataAdapter(query, connection);
                    break;
                case 8:
                    table9 = new OleDbDataAdapter(query, connection);
                    break;
                case 9:
                    table10 = new OleDbDataAdapter(query, connection);
                    break;
                case 10:
                    table11 = new OleDbDataAdapter(query, connection);
                    break;
                case 11:
                    table12 = new OleDbDataAdapter(query, connection);
                    break;
                case 12:
                    table13 = new OleDbDataAdapter(query, connection);
                    break;
                case 13:
                    table14 = new OleDbDataAdapter(query, connection);
                    break;
            }
            RefreshTables();
        }
        static bool isAscending = true;
        private void sortButton_Click(object sender, EventArgs e)
        {
            //Variables

            int tabIndex = tabControl1.SelectedIndex;
            string column = "";

            switch (tabIndex)
            {
                case 0:
                    column = dataGridView1.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 1:
                    column = dataGridView2.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 2:
                    column = dataGridView3.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 3:
                    column = dataGridView4.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 4:
                    column = dataGridView5.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 5:
                    column = dataGridView6.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 6:
                    column = dataGridView7.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 7:
                    column = dataGridView8.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 8:
                    column = dataGridView9.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 9:
                    column = dataGridView10.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 10:
                    column = dataGridView11.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 11:
                    column = dataGridView12.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 12:
                    column = dataGridView13.CurrentCell.OwningColumn.HeaderText;
                    break;
                case 13:
                    column = dataGridView14.CurrentCell.OwningColumn.HeaderText;
                    break;

            }
            if (isAscending == true)
            {
                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text + "] ORDER BY [" + column + "] ASC;";
                isAscending = false;
            }
            else
            {
                query = "SELECT * FROM [" + tabControl1.SelectedTab.Text + "] ORDER BY [" + column + "] DESC;";
                isAscending = true;
            }
            switch (tabIndex)

            {
                case 0:
                    table1 = new OleDbDataAdapter(query, connection);
                    break;
                case 1:
                    table2 = new OleDbDataAdapter(query, connection);
                    break;
                case 2:
                    table3 = new OleDbDataAdapter(query, connection);
                    break;
                case 3:
                    table4 = new OleDbDataAdapter(query, connection);
                    break;
                case 4:
                    table5 = new OleDbDataAdapter(query, connection);
                    break;
                case 5:
                    table6 = new OleDbDataAdapter(query, connection);
                    break;
                case 6:
                    table7 = new OleDbDataAdapter(query, connection);
                    break;
                case 7:
                    table8 = new OleDbDataAdapter(query, connection);
                    break;
                case 8:
                    table9 = new OleDbDataAdapter(query, connection);
                    break;
                case 9:
                    table10 = new OleDbDataAdapter(query, connection);
                    break;
                case 10:
                    table11 = new OleDbDataAdapter(query, connection);
                    break;
                case 11:
                    table12 = new OleDbDataAdapter(query, connection);
                    break;
                case 12:
                    table13 = new OleDbDataAdapter(query, connection);
                    break;
                case 13:
                    table14 = new OleDbDataAdapter(query, connection);
                    break;

            }

            RefreshTables();

        }
        void ResetColor()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {

                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {

                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView3.RowCount; i++)
            {

                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView4.RowCount; i++)
            {

                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView5.RowCount; i++)
            {

                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView6.RowCount; i++)
            {

                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView7.RowCount; i++)
            {

                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView8.RowCount; i++)
            {

                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView9.RowCount; i++)
            {

                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView10.RowCount; i++)
            {

                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView11.RowCount; i++)
            {

                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView12.RowCount; i++)
            {

                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView13.RowCount; i++)
            {

                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
            for (int i = 0; i < dataGridView14.RowCount; i++)
            {

                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void colorizeButton_Click(object sender, EventArgs e)
        {
            ResetColor();
            //Variables
            int tabIndex = tabControl1.SelectedIndex;
            int j = 0;
            Type ColumnValueType = null;
            string FilterValue = "";
            string filterSign = "=";
            bool isString = true;
            //Prompt Form
            Form promptFilter = new Form();
            promptFilter.FormBorderStyle = FormBorderStyle.FixedDialog;
            promptFilter.StartPosition = FormStartPosition.CenterScreen;
            promptFilter.MaximizeBox = false;
            promptFilter.MinimizeBox = false;
            promptFilter.Width = 500;
            promptFilter.Height = 250;
            promptFilter.Text = "Filter";
            RadioButton BiggerButton = new RadioButton() { Text = "Bigger", Left = promptFilter.Width / 2 - 175, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            RadioButton SmallerButton = new RadioButton() { Text = "Smaller", Left = promptFilter.Width / 2 - 50, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            RadioButton EqualsButton = new RadioButton() { Text = "Equals", Left = promptFilter.Width / 2 + 75, Top = 50, Name = "FilterChoice", BackColor = System.Drawing.Color.Transparent };
            Button confirm = new Button() { Text = "Confirm", Left = promptFilter.Width / 2 - 150, Width = 100, Top = promptFilter.Height - 100, Height = 50, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Cancel", Left = promptFilter.Width / 2 + 50, Width = 100, Top = promptFilter.Height - 100, Height = 50, DialogResult = DialogResult.Cancel };
            Label promptLabel = new Label() { Text = "Input Value", Left = promptFilter.Width / 2 - 150, Top = promptFilter.Height - 175, BackColor = System.Drawing.Color.Transparent };
            TextBox promptText = new TextBox() { Left = promptFilter.Width / 2 - 150, Top = promptFilter.Height - 150, Width = promptFilter.Width / 2 };
            promptFilter.Controls.Add(confirm);
            promptFilter.Controls.Add(cancel);
            promptFilter.Controls.Add(promptLabel);
            promptFilter.Controls.Add(promptText);
            cancel.Anchor = AnchorStyles.Bottom;
            confirm.Anchor = AnchorStyles.Bottom;

            switch (tabIndex)
            {
                case 0:
                    ColumnValueType = dataGridView1.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView1.CurrentCell.OwningColumn.Index;
                    break;
                case 1:
                    ColumnValueType = dataGridView2.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView2.CurrentCell.OwningColumn.Index;
                    break;
                case 2:
                    ColumnValueType = dataGridView3.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView3.CurrentCell.OwningColumn.Index;
                    break;
                case 3:
                    ColumnValueType = dataGridView4.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView4.CurrentCell.OwningColumn.Index;
                    break;
                case 4:
                    ColumnValueType = dataGridView5.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView5.CurrentCell.OwningColumn.Index;
                    break;
                case 5:
                    ColumnValueType = dataGridView6.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView6.CurrentCell.OwningColumn.Index;
                    break;
                case 6:
                    ColumnValueType = dataGridView7.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView7.CurrentCell.OwningColumn.Index;
                    break;
                case 7:
                    ColumnValueType = dataGridView8.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView8.CurrentCell.OwningColumn.Index;
                    break;
                case 8:
                    ColumnValueType = dataGridView9.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView9.CurrentCell.OwningColumn.Index;
                    break;
                case 9:
                    ColumnValueType = dataGridView10.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView10.CurrentCell.OwningColumn.Index;
                    break;
                case 10:
                    ColumnValueType = dataGridView11.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView11.CurrentCell.OwningColumn.Index;
                    break;
                case 11:
                    ColumnValueType = dataGridView12.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView12.CurrentCell.OwningColumn.Index;
                    break;
                case 12:
                    ColumnValueType = dataGridView13.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView13.CurrentCell.OwningColumn.Index;
                    break;
                case 13:
                    ColumnValueType = dataGridView14.CurrentCell.OwningColumn.ValueType;
                    j = dataGridView9.CurrentCell.OwningColumn.Index;
                    break;

            }

            if (ColumnValueType != Type.GetType("System.String"))
            {
                promptFilter.Controls.Add(BiggerButton);
                promptFilter.Controls.Add(SmallerButton);
                promptFilter.Controls.Add(EqualsButton);

            }
            promptFilter.ShowDialog();
            if (promptFilter.DialogResult != DialogResult.OK)
            {
                promptFilter.Close();
                return;
            }
            FilterValue = promptText.Text;
            if (int.TryParse(FilterValue, out _))
            {
                isString = false;
            }
            if (ColumnValueType != Type.GetType("System.String"))
            {
                if (BiggerButton.Checked)
                {
                    filterSign = ">";

                }
                if (SmallerButton.Checked)
                {
                    filterSign = "<";

                }
                if (EqualsButton.Checked)
                {
                    filterSign = "=";
                }
            }
            switch (tabIndex)
            {
                case 0:

                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView1.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView1.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView1.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView1.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView1.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 1:
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView2.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView2.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView2.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView2.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView2.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView2.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView2.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 2:
                    for (int i = 0; i < dataGridView3.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView3.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView3.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView3.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView3.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView3.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView3.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView3.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 3:
                    for (int i = 0; i < dataGridView4.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView4.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView4.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView4.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView4.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView4.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView4.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView4.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView4.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView4.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView4.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 4:
                    for (int i = 0; i < dataGridView5.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView5.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView5.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView5.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView5.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView5.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView5.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView5.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView5.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView5.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView5.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 5:
                    for (int i = 0; i < dataGridView6.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView6.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView6.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView6.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView6.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView6.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView6.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView6.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView6.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView6.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView6.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 6:
                    for (int i = 0; i < dataGridView7.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView7.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView7.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView7.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView7.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView7.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView7.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView7.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView7.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView7.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView7.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 7:
                    for (int i = 0; i < dataGridView8.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView8.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView8.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView8.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView8.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView8.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView8.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView8.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView8.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView8.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }

                    break;
                case 8:
                    for (int i = 0; i < dataGridView9.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView9.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView9.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView9.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView9.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView9.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView9.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView9.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView9.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView9.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView9.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 9:
                    for (int i = 0; i < dataGridView10.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView10.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView10.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView10.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView10.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView10.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView10.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView10.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView10.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView10.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView10.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 10:
                    for (int i = 0; i < dataGridView11.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView11.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView11.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView11.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView11.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView11.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView11.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView11.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView11.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView11.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView11.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 11:
                    for (int i = 0; i < dataGridView12.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView12.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView12.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView12.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView12.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView12.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView12.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView12.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView12.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView12.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView12.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 12:
                    for (int i = 0; i < dataGridView13.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView13.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView13.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView13.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView13.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView13.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView13.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView13.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView13.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView13.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView13.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
                case 13:
                    for (int i = 0; i < dataGridView14.RowCount; i++)
                    {
                        if (filterSign == ">")
                        {
                            if (dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView14.Rows[i].Cells[j].Value.ToString()) > int.Parse(FilterValue))
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView14.Rows[i].Cells[j].Value.ToString()).Month > int.Parse(FilterValue))
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "<")
                        {
                            if (dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView14.Rows[i].Cells[j].Value.ToString()) < int.Parse(FilterValue))
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView14.Rows[i].Cells[j].Value.ToString()).Month < int.Parse(FilterValue))
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                            }

                        }
                        if (filterSign == "=")
                        {
                            if (dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.String") && dataGridView14.Rows[i].Cells[j].Value.ToString().ToLower().Contains(FilterValue.ToLower()) && isString == true)
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkViolet;
                            }

                            if (isString == false && dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.Int32") && Convert.ToInt32(dataGridView14.Rows[i].Cells[j].Value.ToString()) == int.Parse(FilterValue))
                            {
                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Crimson;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.CornflowerBlue;
                            }
                            if (isString == false && dataGridView14.Rows[i].Cells[j].ValueType == Type.GetType("System.DateTime") && Convert.ToDateTime(dataGridView14.Rows[i].Cells[j].Value.ToString()).Month == int.Parse(FilterValue))
                            {

                                dataGridView14.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Gray;
                                dataGridView14.Rows[i].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;

                            }
                        }
                    }
                    break;
            }


        }

        private void printButton_Click(object sender, EventArgs e)
        {
            try
            {
                string TransactionID = dataGridView8.SelectedRows[0].Cells[0].Value.ToString();
                query = "SELECT account.user_account_name, account.user_account_login, store_item.store_item_name, store_item.store_item_price, transaction_dot FROM store_item INNER JOIN(account INNER JOIN transactions ON account.user_account_login = transactions.user_account_login) ON store_item.store_item_id = transactions.store_item_id WHERE transactions.transaction_id = " + TransactionID + ";";
                command = new OleDbCommand(query, connection);
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                dataTable.Rows.Clear();
                adapter.Fill(dataTable);
                string dataDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Document document = new Document();
                Page page = document.Pages.Add();
                Aspose.Pdf.Text.TextFragment companyNameLogo = new Aspose.Pdf.Text.TextFragment("Bored Bakers");
                companyNameLogo.IsInLineParagraph = true;
                companyNameLogo.VerticalAlignment = VerticalAlignment.Bottom;
                companyNameLogo.TextState.FontSize = 20;
                companyNameLogo.Margin = new MarginInfo(0, 0, 0, 75);
                Aspose.Pdf.Image image = new Aspose.Pdf.Image();
                image.File = dataDir + "//logo.png";
                image.FixHeight = 100;
                image.FixWidth = 100;
                image.IsInLineParagraph = true;
                image.VerticalAlignment = VerticalAlignment.Bottom;
                page.Paragraphs.Add(image);
                page.Paragraphs.Add(companyNameLogo);
                Aspose.Pdf.Text.TextFragment GreetingsText = new Aspose.Pdf.Text.TextFragment();
                GreetingsText.Text = "Greetings, " + dataTable.Rows[0].ItemArray[0].ToString();
                GreetingsText.Margin = new MarginInfo(0, 0, 0, 15);
                GreetingsText.TextState.FontSize = 16;
                page.Paragraphs.Add(GreetingsText);
                Aspose.Pdf.Text.TextFragment AppreciationText = new Aspose.Pdf.Text.TextFragment();
                AppreciationText.Text = "Thank you for purchasing an item on the store.";
                AppreciationText.Margin = new MarginInfo(0, 0, 0, 15);
                AppreciationText.TextState.FontSize = 16;

                page.Paragraphs.Add(AppreciationText);
                Aspose.Pdf.Text.TextFragment ItemText = new Aspose.Pdf.Text.TextFragment();
                ItemText.Text = "The item below will appear in your account's inventory soon.";
                ItemText.Margin = new MarginInfo(0, 0, 0, 15);
                ItemText.TextState.FontSize = 16;
                page.Paragraphs.Add(ItemText);

                Aspose.Pdf.Text.TextFragment AccountName = new Aspose.Pdf.Text.TextFragment();
                AccountName.Text = "Account Name: " + dataTable.Rows[0].ItemArray[1].ToString();
                AccountName.Margin = new MarginInfo(0, 0, 0, 30);
                AccountName.TextState.FontSize = 12;
                page.Paragraphs.Add(AccountName);

                Aspose.Pdf.Text.TextFragment BoughtItemName = new Aspose.Pdf.Text.TextFragment();
                BoughtItemName.Text = "Item Name";
                BoughtItemName.Margin = new MarginInfo(0, 0, 0, 30);
                BoughtItemName.TextState.FontSize = 12;
                page.Paragraphs.Add(BoughtItemName);

                Aspose.Pdf.Text.TextFragment BoughtItemPrice = new Aspose.Pdf.Text.TextFragment();
                BoughtItemPrice.Text = "Item Price";
                BoughtItemPrice.Margin = new MarginInfo(250, 0, 0, 0);
                BoughtItemPrice.TextState.FontSize = 12;
                BoughtItemPrice.IsInLineParagraph = true;
                page.Paragraphs.Add(BoughtItemPrice);

                Aspose.Pdf.Text.TextFragment ItemName = new Aspose.Pdf.Text.TextFragment();
                ItemName.Text = dataTable.Rows[0].ItemArray[2].ToString();
                ItemName.Margin = new MarginInfo(0, 0, 0, 30);
                ItemName.TextState.FontSize = 12;
                page.Paragraphs.Add(ItemName);

                Aspose.Pdf.Text.TextFragment ItemPrice = new Aspose.Pdf.Text.TextFragment();
                ItemPrice.Text = dataTable.Rows[0].ItemArray[3].ToString() + "$";
                ItemPrice.Margin = new MarginInfo(250, 0, 0, 0);
                ItemPrice.TextState.FontSize = 12;
                ItemPrice.IsInLineParagraph = true;
                page.Paragraphs.Add(ItemPrice);

                string date = dataTable.Rows[0].ItemArray[4].ToString();
                date = date.Remove(date.Length - 12);
                Aspose.Pdf.Text.TextFragment Date = new Aspose.Pdf.Text.TextFragment();
                Date.Text = "Date of Transaction: " + date;
                Date.Margin = new MarginInfo(0, 0, 0, 30);
                Date.TextState.FontSize = 12;
                page.Paragraphs.Add(Date);

                document.Save(dataDir + "//Receipt.pdf");
                //Opens the pdf file
                System.Diagnostics.Process.Start(dataDir + "//Receipt.pdf");
            }
            catch
            {
                MessageBox.Show("Error: You didn't choose the row");
            }


        }
        private void printQueue_Click(object sender, EventArgs e)
        {
            //Variables
            string dataDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Document document = new Document();
            Page page = document.Pages.Add();
            DataGridView datagrid = dataGridView15;
            Aspose.Pdf.Table table = new Aspose.Pdf.Table();
            Aspose.Pdf.Text.TextFragment CreationTime = new Aspose.Pdf.Text.TextFragment();
            Aspose.Pdf.Table NumberOfRecords = new Aspose.Pdf.Table();
            Aspose.Pdf.Text.TextFragment TableName = new Aspose.Pdf.Text.TextFragment();

            page.PageInfo.Margin.Left = 30;
            page.PageInfo.Margin.Right = 30;
            //Report generation time
            CreationTime.Text = "Report generated on " + System.DateTime.Now.ToString();
            CreationTime.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Right;
            //Name of the table
            TableName.Text = "Accounts Characters";
            TableName.HorizontalAlignment = Aspose.Pdf.HorizontalAlignment.Center;
            TableName.TextState.FontSize = 15;
            TableName.TextState.FontStyle = Aspose.Pdf.Text.FontStyles.Bold;
            //Table design
            table.ColumnAdjustment = Aspose.Pdf.ColumnAdjustment.AutoFitToWindow;
            table.DefaultCellPadding = new MarginInfo() { Bottom = 10, Left = 5, Right = 5, Top = 10 };
            table.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.WhiteSmoke));
            table.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.DarkRed));
            //Table population
            Aspose.Pdf.Row HeaderRow = table.Rows.Add();
            for (int col = 0; col < datagrid.Columns.Count; col++)
            {
                HeaderRow.Cells.Add(datagrid.Columns[col].HeaderText);
                HeaderRow.Cells[col].Alignment = Aspose.Pdf.HorizontalAlignment.Center;
                HeaderRow.BackgroundColor = Aspose.Pdf.Color.WhiteSmoke;
            }
            int i = 0;
            for (i = 0; i < datagrid.Rows.Count; i++)
            {
                Aspose.Pdf.Row Row = table.Rows.Add();
                for (int col = 0; col < datagrid.Columns.Count; col++)
                {
                    Row.Cells.Add(datagrid.Rows[i].Cells[col].Value.ToString());
                    Row.Cells[col].Alignment = Aspose.Pdf.HorizontalAlignment.Center;
                    Row.BackgroundColor = Aspose.Pdf.Color.Azure;
                }
            }
            //Table with number count of records in a previous table
            NumberOfRecords.DefaultCellPadding = new MarginInfo() { Bottom = 10, Left = 5, Right = 5, Top = 10 };
            Aspose.Pdf.Row records = NumberOfRecords.Rows.Add();
            records.Cells.Add("Number of Records");
            records.Cells[0].BackgroundColor = Aspose.Pdf.Color.LightCyan;
            records.Cells[0].Alignment = Aspose.Pdf.HorizontalAlignment.Center;
            records = NumberOfRecords.Rows.Add();
            records.Cells.Add(i.ToString());
            records.Cells[0].BackgroundColor = Aspose.Pdf.Color.LightGreen;
            records.Cells[0].Alignment = Aspose.Pdf.HorizontalAlignment.Center;
            //Adding elements to the page
            page.Paragraphs.Add(CreationTime);
            page.Paragraphs.Add(TableName);
            page.Paragraphs.Add(table);
            page.Paragraphs.Add(NumberOfRecords);
            document.Save(dataDir + "//Report.pdf");
            //Opens the pdf file
            System.Diagnostics.Process.Start(dataDir + "//Report.pdf");

        }
        private void resetButton_Click(object sender, EventArgs e)
        {
            SetTables();
            RefreshTables();
        }

    }
}

