using System.Collections;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.ComponentModel.Design;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;

namespace KR_DailyReport
{

    public partial class FormMain : Form
    {
        // DB�ڑ�������
        const string CONNECT_STRING = "Database=KRDairyDB;Integrated Security=SSPI;Persist Security Info=False;Connection Timeout=10";
        //const string CONNECT_STRING = "Integrated Security=SSPI;Persist Security Info=False;Connection Timeout=10";

        // �ϐ�
        public NET10Address? m_NET10Addr = null; // ���^�A�h���X�Ǘ��N���X
        public NET10Control? m_NET10Ctrl = null; // NET10�ʐM�Ǘ��N���X

        // NET10�A�h���X���(Excel�t�@�C���Ǐo�p)
        public struct ExcelInfo
        {
            public List<int> Row;       // Excel�V�[�g�̍s�ԍ�
            public List<string> Item;   // ���ږ�
            public List<short> Addr;    // �A�h���X
            public List<short> Data;    // �����f�[�^
        }
        private string m_TemplateName = Properties.Settings.Default.TemplateName; // �ݒ�t�@�C������擾
        private string m_SavePath = Properties.Settings.Default.SavePath; // �ݒ�t�@�C������擾
        private ExcelInfo[] m_Excel = new ExcelInfo[2]; // Excel���
        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Microsoft.Office.Interop.Excel.Workbook? xlWorkbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet? xlWorksheet = null;

        private SqlConnection m_SQLConnect; // SQL�ڑ��N���X�iADO.NET�j

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public FormMain()
        {
            InitializeComponent();

            m_Excel[0] = new ExcelInfo();
            m_Excel[1] = new ExcelInfo();
        }

        /// <summary>
        /// �f�X�g���N�^
        /// </summary>
        ~FormMain()
        {
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }

            if (xlWorkbook != null)
            {
                // Excel Close and Release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }

            // Excel Quit and Release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        /// <summary>
        /// �t�H�[�����[�h
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;

#if false
            ////////////////////////////////////////
            // NET10�ʐM�̏�����
            m_NET10Ctrl = new NET10Control();
            if (m_NET10Ctrl.StartComm() == false)
            {
                MessageBox.Show("NET10�ʐM�̏������Ɏ��s���܂����B", "NET10�ʐM�������G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
#endif
            dgvDailyData.ColumnCount = 2;

            dgvDailyData.Columns[0].Width = 270;
            dgvDailyData.Columns[0].HeaderText = "���ږ�";

            dgvDailyData.Columns[1].Width = 150;
            dgvDailyData.Columns[1].HeaderText = "�l";

            dtpDate.Value = DateTime.Now;
        }

        /// <summary>
        /// �p�����[�^�̃R���{���X�g���쐬����
        /// </summary>
        private void MakeParamCombo()
        {
            lblStatus.Text = "Excel�t�@�C���Ǎ���...";

            // �p�����[�^���A�A�h���X�A�ݒ�l�AExcel�̃Z���ʒu
            //ExcelRead(m_ExcelPath, 1, 2, 2, 200, 5);

            lblStatus.Text = "";
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
#if false

            if (m_NET10Ctrl != null)
            {
                if (m_NET10Ctrl.EndComm() == false)
                {
                    MessageBox.Show("NET10�̒ʐM��~�Ɏ��s���܂����B", "NET10�ʐM��~�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
#endif
        }

        /// <summary>
        /// Excel�֓���o��
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcelOutput_Click(object sender, EventArgs e)
        {
            if (dgvDailyData.Rows .Count == 0)
            {
                MessageBox.Show("�f�[�^�����o����Ă��܂���", "�f�[�^���I��", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Excel�֓�����o�͂��܂��B��낵���ł����H\n" + "(�ۑ���F" + m_SavePath + ")", "����o��", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {

                var data = new ArrayList();
                for (int i = 0; i < dgvDailyData.RowCount; i++)
                {
                    var cell = dgvDailyData[1, i].Value;

                    if (cell == null || cell.ToString() == "")
                    {
                        data.Add("");
                    }
                    else
                    {
                        data.Add(dgvDailyData[1, i].Value.ToString());
                    }
                }

                ExcelWrite(m_SavePath, 1, 3, 4, data, dtpDate.Value.ToString());

                MessageBox.Show("����o�͂��������܂����B", "�����m�F", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string MakeSavePath()
        {
            //SaveFileDialog�N���X�̃C���X�^���X���쐬
            SaveFileDialog sfd = new SaveFileDialog();

            //�͂��߂̃t�@�C�������w�肷��
            //�͂��߂Ɂu�t�@�C�����v�ŕ\������镶������w�肷��
            sfd.FileName = "����.xlsx";
            //�͂��߂ɕ\�������t�H���_���w�肷��
            sfd.InitialDirectory = m_SavePath;
            //[�t�@�C���̎��]�ɕ\�������I�������w�肷��
            //�w�肵�Ȃ��i��̕�����j�̎��́A���݂̃f�B���N�g�����\�������
            sfd.Filter = "Excel�t�@�C��(*.xlsx)|*.xlsx|���ׂẴt�@�C��(*.*)|*.*";
            //[�t�@�C���̎��]�ł͂��߂ɑI���������̂��w�肷��
            //2�Ԗڂ́u���ׂẴt�@�C���v���I������Ă���悤�ɂ���
            sfd.FilterIndex = 2;
            //�^�C�g����ݒ肷��
            sfd.Title = "�ۑ���̃t�@�C����I�����Ă�������";
            //�_�C�A���O�{�b�N�X�����O�Ɍ��݂̃f�B���N�g���𕜌�����悤�ɂ���
            sfd.RestoreDirectory = true;
            //���ɑ��݂���t�@�C�������w�肵���Ƃ��x������
            //�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
            sfd.OverwritePrompt = true;
            //���݂��Ȃ��p�X���w�肳�ꂽ�Ƃ��x����\������
            //�f�t�H���g��True�Ȃ̂Ŏw�肷��K�v�͂Ȃ�
            sfd.CheckPathExists = true;

            //�_�C�A���O��\������
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //OK�{�^�����N���b�N���ꂽ�Ƃ��A�I�����ꂽ�t�@�C������\������
                Console.WriteLine(sfd.FileName);
            }

            return sfd.FileName;
        }

#if false
        /// <summary>
        /// �G�N�Z���t�@�C���̎w�肵���V�[�g��2�����z��ɓǂݍ���.
        /// </summary>
        /// <param name="filePath">�G�N�Z���t�@�C���̃p�X</param>
        /// <param name="sheetIndex">�V�[�g�̔ԍ� (1, 2, 3, ...)</param>
        /// <param name="startRow">�ŏ��̍s (>= 1)</param>
        /// <param name="startColmn">�ŏ��̗� (>= 1)</param>
        /// <param name="lastRow">�Ō�̍s</param>
        /// <param name="lastColmn">�Ō�̗�</param>
        /// <returns>�V�[�g�����i�[����2���������z��. �������t�@�C���ǂݍ��݂Ɏ��s�����Ƃ��ɂ� null.</returns>
        public ArrayList ExcelRead(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            var arrOut = new ArrayList();

            m_Excel[sheetIndex - 1].Row = new List<int>();
            m_Excel[sheetIndex - 1].Item = new List<string>();
            m_Excel[sheetIndex - 1].Addr = new List<short>();
            m_Excel[sheetIndex - 1].Data = new List<short>();

            // ���[�N�u�b�N���J��
            if (!ExcelOpen(filePath)) { return arrOut; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();


                for (int r = startRow; r <= lastRow; r++)
                {
                    // ��s�ǂݍ���
                    var row = new ArrayList();
                    for (int c = startColmn; c <= lastColmn; c++)
                    {
                        var cell = xlWorksheet.Cells[r, c];

                        if (cell == null || cell.Value == null)
                        {
                            row.Add("");
                        }
                        else
                        {
                            row.Add(cell.Value);
                        }
                    }

                    string item = "";
                    short addr = 0;

                    // ���ږ�
                    if (row[0] != null && row[0] != "")
                    {
                        item = row[0].ToString();
                    }
                    else
                    {
                        // �Ǎ��I��
                        break;
                    }

                    // �A�h���X
                    if (row[3] != null)
                    {
                        if (row[3].ToString() != "")
                        {
                            addr = Convert.ToInt16(row[3].ToString().Substring(1, 4), 16);
                        }
                    }

                    short data = 0;
                    if (sheetIndex == 1)
                    {
                        // ��ƃf�[�^
                        if (row[2] != null)
                        {
                            if (row[2].ToString() != "")
                            {
                                data = Convert.ToInt16(row[2].ToString(), 10);
                            }
                        }
                    }
                    else
                    {
                        // ����f�[�^
                        if (row[2] != null)
                        {
                            if (row[2].ToString() != "")
                            {
                                data = Convert.ToInt16(row[2].ToString(), 10);
                            }
                        }

                    }

                    //���X�g�Ɋi�[
                    m_Excel[sheetIndex - 1].Row.Add(r);
                    m_Excel[sheetIndex - 1].Item.Add(item);
                    m_Excel[sheetIndex - 1].Addr.Add(addr);
                    m_Excel[sheetIndex - 1].Data.Add(data);

                    arrOut.Add(row);
                }

                // ���[�N�V�[�g�����
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
            ExcelClose();

            return arrOut;
        }
#endif

        /// <summary>
        /// �G�N�Z���t�@�C���̎w�肵���V�[�g�̃Z���Ƀf�[�^����������
        /// </summary>
        /// <param name="filePath">�G�N�Z���t�@�C���̃p�X</param>
        /// <param name="sheetIndex">�V�[�g�̔ԍ� (1, 2, 3, ...)</param>
        /// <param name="row">�s (>= 1)</param>
        /// <param name="colmn">�� (>= 1)</param>
        /// <param name="data">�����f�[�^</param>
        /// <returns>���s����</returns>
        public bool ExcelWrite(string filePath, int sheetIndex, int row, int colmn, ArrayList data, string date)
        {
            this.Cursor = Cursors.WaitCursor;

            lblStatus.Text = "Excel�t�@�C���֕ۑ���...";

            // �e���v���[�g���R�s�[
            if (!File.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            // �f�[�^�擾�N����
            string strTime = dtpDate.Value.Date.ToString().Substring(0,11) + cmbStartTimeStamp.SelectedItem.ToString(); ;
            DateTime dTime = DateTime.Parse(strTime);

            string SaveFilePath = filePath + dTime.ToString("����yyyyMMdd_HHmmss") + ".xlsx";
            File.Copy(AppDomain.CurrentDomain.BaseDirectory + @"Template\" + m_TemplateName, SaveFilePath);

            // ���[�N�u�b�N���J��
            if (!ExcelOpen(SaveFilePath)) { return false; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();

                xlWorksheet.Cells[1, 4] = date.Substring(0, 10) + "(1/1)";     // ���t

                // �f�[�^
                xlWorksheet.Cells[row, colmn] = data[data.Count - 2];
                xlWorksheet.Cells[row + 1, colmn] = data[data.Count - 1];

                for (int i = 0; i < data.Count - 2; i++)
                {
                    xlWorksheet.Cells[i + row + 2, colmn] = data[i];
                }

                // �ۑ�����
                xlWorkbook.Save();

                // ���[�N�V�[�g�����
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
            ExcelClose();

            lblStatus.Text = "";

            this.Cursor = Cursors.Default;

            return true;
        }

        /// <summary>
        /// �w�肳�ꂽ�p�X��Excel Workbook���J��
        /// </summary>
        /// <param name="filePath">Excel Workbook�̃p�X(���΃p�X�ł���΃p�X�ł�OK)</param>
        /// <returns>Excel Workbook�̃I�[�v���ɐ��������� true. ����ȊO false.</returns>
        protected bool ExcelOpen(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.Visible = false;

                // filePath �����΃p�X�̂Ƃ���O����������̂� fullPath �ɕϊ�
                string fullPath = System.IO.Path.GetFullPath(filePath);
                xlWorkbook = xlApp.Workbooks.Open(fullPath);
            }
            catch
            {
                ExcelClose();
                return false;
            }

            return true;
        }

        /// <summary>
        /// �J���Ă���Workbook��Excel�̃v���Z�X�����.
        /// </summary>
        protected void ExcelClose()
        {
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }

            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }

        /// <summary>
        /// �p�����[�^�ύX�i����p�����[�^�j
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// ����p�����[�^�̓��͒l��Net10�̎w��A�h���X�ɑ��M
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet1_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// ���t�ύX��
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dtpDate_ValueChanged(object sender, EventArgs e)
        {
            TimeList();
        }

        /// <summary>
        /// �Ώۓ��t�̎��������X�g��
        /// </summary>
        private void TimeList()
        {
            SqlConnection con = new SqlConnection(CONNECT_STRING);
            con.Open();

            try
            {
                int i = 0;
                var strList = new List<string>();

                cmbStartTimeStamp.Items.Clear();
                cmbStartTimeStamp.Text = "";
                string st_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day + " 00:00:00";
                string ed_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day + " 23:59:59";
                string sqlstr = "SELECT �J�n���� FROM JOB WHERE �J�n���� BETWEEN '" + st_date + "' AND '" + ed_date + "'";
                SqlCommand com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        if (reader[0] != null)
                        {
                            cmbStartTimeStamp.Items.Add(reader[0].ToString().Substring(11));
                        }
                        i++;
                    }
                }
                com.Dispose();
            }
            finally
            {
                con.Close();
            }

            if (cmbStartTimeStamp.Items.Count > 0)
            {
                cmbStartTimeStamp.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// �f�[�^���o
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetData_Click(object sender, EventArgs e)
        {
            if (dtpDate.Text.ToString() == "")
            {
                MessageBox.Show("���t�����͂���Ă��܂���", "���̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbStartTimeStamp.SelectedItem == null || cmbStartTimeStamp.SelectedItem.ToString() == "")
            {
                MessageBox.Show("�J�n�������I������Ă��܂���", "���̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SqlConnection con = new SqlConnection(CONNECT_STRING);
            con.Open();

            try
            {
                var strHeadList = new List<string>();
                var strList = new List<string>();

                string sqlstr = "SELECT t.name AS �e�[�u���� ,c.name AS ���ږ� ,type_name(user_type_id) AS ���� , max_length AS ���� , CASE WHEN is_nullable = 1 THEN 'YES' ELSE 'NO' END AS NULL���� FROM sys.objects t INNER JOIN sys.columns c ON t.object_id = c.object_id WHERE t.type = 'U' AND t.name = 'JOB' ORDER BY c.column_id";
                SqlCommand com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        strHeadList.Add(reader[1].ToString());
                    }
                }

                string tgt_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day;
                sqlstr = "SELECT * FROM JOB WHERE �J�n���� ='" + tgt_date + " " + cmbStartTimeStamp.SelectedItem + "'";
                com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (reader[i] != null)
                            {
                                strList.Add(reader[i].ToString());
                            }
                            else
                            {
                                strList.Add("");
                            }
                        }
                    }
                }

                dgvDailyData.Rows.Clear();
                for (int i = 1; i < strHeadList.Count; i++)
                {
                    dgvDailyData.Rows.Add(strHeadList[i], strList[i]);
                }
                com.Dispose();

            }
            finally
            {
                con.Close();
            }
        }
    }
}