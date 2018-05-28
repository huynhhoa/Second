using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace Second
{
    public partial class CVExcel : Form
    {
        public CVExcel()
        {
            InitializeComponent();
        }
        Workbook book;
        Worksheet sheet;
        private void btnPF_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "|*.xlsx";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                try
                {
                    dgv1.DataSource = null; //làm sạch dgv1
                    dgv1.Rows.Clear();
                    txtPF.Text = openFileDialog.FileName;
                    book = new Workbook();
                    book.LoadFromFile(txtPF.Text);
                    sheet = book.Worksheets[0];
                    string[] splitText = null;
                    string text = null;
                    DateTime dt;
                    //xử lý Sheet
                    #region 10
                  //tách hàng
                    for (int i = 1; i < sheet.LastRow; i++)
                    {
                        text = sheet.Range[i + 1, 10].Text; //cột thứ 10
                    //nếu dòng đó null thì bỏ
                        if (text != null)
                        {
                            int k = i + 1;
                            splitText = text.Split('\n');
                            if (splitText.Length - 1 > 0 || (splitText.Length == 1 && i == sheet.LastRow - 1)) //Nếu độ dài của mảng =2 thì thực hiện. 
                            {
                               // int k = i + 1;
                                for (int j = 0; j < splitText.Length; j++)
                                {
                                    sheet.InsertRow(k + 1); //chèn thêm 1 hàng
                                    sheet.Range[k, 10].Text = splitText[j];
                                    sheet.Range[k, 3].Value = sheet.Range[i + 1, 3].Value;
                                    sheet.Range[k, 9].Value = sheet.Range[i + 1, 9].Value;
                                    sheet.Range[k, 6].Value = sheet.Range[i + 1, 6].Value;
                                    k++;
                                    if (j == splitText.Length - 1) //xóa dòng cuối (dòng trống) 
                                    {
                                        sheet.DeleteRow(k);
                                    }
                                }
                            }
                            
                        }
                        

                    }
                 
                    // tách số trong hàng
                    for (int i = 1; i < sheet.LastRow; i++)
                    {
                        text = sheet.Range[i + 1, 10].Text; //text bằng cột thứ 10, dòng 2
                        if (text != null && text.Contains("use") == false)
                        {
                            splitText = text.Split('(');
                            int k = i + 1;
                            for (int j = 1; j < splitText.Length; j++)
                            {
                                sheet.Range[k, 10].Text = splitText[j];
                                k++;
                            }

                        }
                        else if (text != null && text.Contains("use"))
                        {
                            splitText = text.Split('e');
                            int k = i + 1;
                            for (int j = 1; j < splitText.Length; j++)
                            {
                                sheet.Range[k, 10].Text = splitText[j];
                                k++;
                            }
                        }
                    }
                   //tách chữ p thường
                    for (int i = 1; i < sheet.LastRow; i++)
                    {

                        text = sheet.Range[i + 1, 10].Text;
                        if (text != null)
                        {
                            splitText = text.Split('p');
                            sheet.Range[i + 1, 10].Text = splitText[0];

                        }
                    }
                //tách chữ P HOA
                for (int i = 1; i < sheet.LastRow; i++)
                {

                    text = sheet.Range[i + 1, 10].Text;
                    if (text != null)
                    {
                        splitText = text.Split('P');
                        sheet.Range[i + 1, 10].Text = splitText[0];

                    }
                }
                #endregion
                #region 11
                //duyệt từ hàng thứ 1. 
                for (int i = 1; i < sheet.LastRow; i++)
                    {

                        text = sheet.Range[i + 1, 11].Text; //cột thứ 11
                                                            //nếu dòng đó null thì bỏ
                        if (text != null)
                        {
                            splitText = text.Split('\n');
                            if (splitText.Length - 1 > 0)
                            {
                                int k = i + 1;
                                for (int j = 0; j < splitText.Length; j++)
                                {
                                    sheet.Range[k, 11].Value = splitText[j]; //cột thứ 11
                                    k++;
                                }
                            }
                        }
                    }
                #endregion
                    dgv1.ColumnCount = 15;
                    for (int i = 0; i < sheet.Rows.Length -1; i++)
                    {
                        dgv1.Rows.Add();
                        dgv1.Rows[i].Cells[0].Value = "NEV";
                        dgv1.Rows[i].Cells[1].Value = "DDKV";
                        dgv1.Rows[i].Cells[2].Value = "VS2";
                        dgv1.Rows[i].Cells[6].Value = "T";
                        dgv1.Rows[i].Cells[7].Value = "F";
                        dgv1.Rows[i].Cells[8].Value = "";
                    dgv1.Rows[i].Cells[9].Value = sheet.Range[i + 1, 9].Value;

                    dgv1.Rows[i].Cells[10].Value = "NEC";
                    dgv1.Rows[i].Cells[11].Value = sheet.Range[i + 1, 3].Value;

                    dgv1.Rows[i].Cells[12].Value = "VN";
                    dgv1.Rows[i].Cells[13].Value = sheet.Range[i + 1, 10].Value.ToString();
                    dgv1.Rows[i].Cells[14].Value = sheet.Range[i + 1, 11].Value.ToString();

                }
             
                //kiểm tra nếu ô đó là định dạng ngày thì ô thứ 5 bằng ô đó
                for (int i = 0; i < dgv1.Rows.Count-1; i++)
                {
                    if (DateTime.TryParse(sheet.Range[i+1, 6].Value, out dt))
                    {
                        dgv1.Rows[i ].Cells[3].Value = DateTime.Parse(sheet.Range[i+1, 6].Value).ToString("yyyyMMdd");
                    }
                    else
                    {
                        dgv1.Rows[i ].Cells[3].Value = dgv1.Rows[i - 1].Cells[3].Value;
                    }
                }
                //cột 4 của dgv1. Nếu không phải ngày thì lấy
                //c2: lấy hết cột 4. nếu là ngày thì xóa
                for (int i = 0; i < sheet.LastRow; i++)
                    {
                       dgv1.Rows[i].Cells[4].Value = sheet.Range[i+1, 6].Value;
                    }
                for (int i = 0; i < dgv1.Rows.Count-1; i++)
                {
                    if (dgv1.Rows[i].Cells[4].Value.ToString() ==""|| DateTime.TryParse(dgv1.Rows[i].Cells[4].Value.ToString(), out dt))
                    {
                        dgv1.Rows.RemoveAt(i);
                    }
                }
                   MessageBox.Show("Conversion Successful", "Sucess", MessageBoxButtons.OK, MessageBoxIcon.Information);
           
            }
                catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (dgv1.Rows.Count != 0)
            {
                try
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "CSV (*.csv)|*.csv";
                    sfd.FileName = "ORDER.csv";
                   
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        

                        // create one file gridview.csv in writing mode using streamwriter
                        StreamWriter sw = new StreamWriter(sfd.FileName);
                        // iterate through all the rows within the gridview
                        foreach (DataGridViewRow dr in dgv1.Rows)
                        {
                            // iterate through all colums of specific row
                            for (int i = 0; i < dgv1.Columns.Count; i++)
                            {
                                // write particular cell to csv file
                                sw.Write(dr.Cells[i].Value);
                                if (i != dgv1.Columns.Count)
                                {
                                    sw.Write(",");
                                }
                            }
                            // write new line
                            sw.Write(sw.NewLine);
                        }
                        // FileInfo fileInfo = new FileInfo(sfd.FileName);
                        FileAttributes file = File.GetAttributes(sfd.FileName);
                        File.SetAttributes(sfd.FileName, FileAttributes.Normal);
                        MessageBox.Show("Save Successfully", "Note",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        // flush from the buffers.
                        sw.Flush();
                    
                        // closes the file
                        sw.Close();
                        
                       // fileInfo.IsReadOnly = false;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("You do not have any data", "NOTE");

            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About ab = new About();
            ab.Show();
        }
    }
}
