using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DeliveryInstructionCreateTool
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.TextBox[] textBoxes;  // クラスのフィールドとして宣言
        private System.Windows.Forms.Button outputButton;
        private Panel scrollPanel;
        private TableLayoutPanel layout;
        private DataGridView dataGridView;
        private DataTable dataTable;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void InitializeCustomComponents()
        {
            // スクロール可能なパネルの作成
            scrollPanel = new Panel
            {
                AutoScroll = true,
                Dock = DockStyle.Top,
                Height = 600 // 高さを調整して表示領域を広げる
            };

            // グリッドレイアウトパネルの作成
            layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            // ヘッダー行の作成
            string[] headers = {
            "SHIPMENT_ID", "STATUS", "DELIVERY_KBN", "KUKURI_NO", "DELIVERY_NAME", "DELIVERY_NAME_KANA",
            "ATENA", "ZIP_CD", "FUKEN_ID", "ADDRESS", "TEL_NO", "BINSYU_ID", "SHIP_DATE", "DELIVERY_DATE",
            "SPECIFY_TIME", "OKURIJYO_NOTE", "DELIVERY_NOTE", "ATTACHMENT_NAME", "MOUSIOKURI_NOTE", "SHIP_COMPANY",
            "STOP_FLG", "USER_ID", "DHC_SALES_ID", "CREATE_STAMP", "ITEM_ID", "ITEM_NAME", "ORDER_NUM", "WMS_ITEM_NUM",
            "WMS_SHIP_DATE", "WMS_DELIVERY_DATE", "SLIP_NUMBER", "CUSTOMER_CD", "CUSTOMER_CD_SUB", "SLIP_NUMBER_HYOZI",
            "REQ_CD", "SHIPPING_NAME", "SHIPPING_YUBIN", "SHIPPING_ADDRESS", "SHIPPING_TEL", "DAIKO_NINUSI_KBN",
            "HATUTEN_CD", "HAISO_TOI_TEL", "CASH_ON_DELIVERY_PRICE", "GIANTSTIME", "FUKEN_NAME", "DEL_TIME_DENPYO",
            "DEL_TIME_DENPYO_JUTAKU", "INZI_BINSYU", "ZYUTAKU_BINSYU", "TOTYAKU_GENPYO_KBNCD", "SLIP_NUMBER_KBN",
            "KESSAI", "RYOSYU_KBN", "HUKA_SERVICE_KBN", "FORMFEED_KBN", "OUTPUT_SN", "INPUT_SN", "ERROR_MSG"
        };

            // テーブルレイアウトパネルの列の幅を設定
            layout.ColumnCount = 2;  // 2列の設定
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200F)); // ラベル用の列
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F)); // テキストボックス用の列

            // 各コントロールを追加
            textBoxes = new TextBox[headers.Length];
            for (int i = 0; i < headers.Length; i++)
            {
                var lbl = new Label
                {
                    Text = headers[i],
                    AutoSize = true,
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };
                var txt = new TextBox { Width = 250 };

                textBoxes[i] = txt;

                // UIに表示するラベルとテキストボックスを追加
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 行の高さを自動調整

                layout.Controls.Add(lbl, 0, i); // ラベルを左列に追加
                layout.Controls.Add(txt, 1, i); // テキストボックスを右列に追加
            }

            scrollPanel.Controls.Add(layout);

            // DataGridViewの設定
            dataTable = new DataTable();
            foreach (string header in headers)
            {
                dataTable.Columns.Add(header);
            }

            dataGridView = new DataGridView
            {
                DataSource = dataTable,
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false
            };

            // 出力ボタンの設定
            outputButton = new Button
            {
                Text = "出力",
                Dock = DockStyle.Bottom
            };
            outputButton.Click += OutputButton_Click;

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };
            buttonPanel.Controls.Add(outputButton);

            var addButton = new Button
            {
                Text = "データ追加",
                Dock = DockStyle.Bottom
            };
            addButton.Click += AddButton_Click;

            buttonPanel.Controls.Add(addButton);

            this.Controls.Add(dataGridView);
            this.Controls.Add(scrollPanel);
            this.Controls.Add(buttonPanel);
        }

        private void OutputButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                FileName = "shipment_instructions.csv"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.GetEncoding("shift_jis")))
                    {
                        // ヘッダー行の出力
                        string[] headers = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName).ToArray();
                        writer.WriteLine(string.Join(",", headers)); // ヘッダー行を出力

                        // データ行の出力
                        foreach (DataRow row in dataTable.Rows)
                        {
                            writer.WriteLine(string.Join(",", row.ItemArray));
                        }
                    }

                    MessageBox.Show("ファイルが正常に保存されました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            var newRow = dataTable.NewRow();
            for (int i = 0; i < textBoxes.Length; i++)
            {
                newRow[i] = textBoxes[i].Text;
            }
            dataTable.Rows.Add(newRow);
        }
    }
}
