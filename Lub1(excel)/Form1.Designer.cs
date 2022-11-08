namespace Lub1_excel_
{
    partial class MyExcel
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.FileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SaveFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SaveFileAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.EditToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddColumToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteRowToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteColToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.clearTableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.informationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.FileToolStripMenuItem,
            this.EditToolStripMenuItem,
            this.informationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(891, 40);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // FileToolStripMenuItem
            // 
            this.FileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newFileToolStripMenuItem,
            this.OpenFileToolStripMenuItem,
            this.SaveFileToolStripMenuItem,
            this.SaveFileAsToolStripMenuItem});
            this.FileToolStripMenuItem.Name = "FileToolStripMenuItem";
            this.FileToolStripMenuItem.Size = new System.Drawing.Size(90, 36);
            this.FileToolStripMenuItem.Text = "Файл";
            // 
            // newFileToolStripMenuItem
            // 
            this.newFileToolStripMenuItem.Name = "newFileToolStripMenuItem";
            this.newFileToolStripMenuItem.Size = new System.Drawing.Size(281, 44);
            this.newFileToolStripMenuItem.Text = "Новий файл";
            this.newFileToolStripMenuItem.Click += new System.EventHandler(this.newFileToolStripMenuItem_Click);
            // 
            // OpenFileToolStripMenuItem
            // 
            this.OpenFileToolStripMenuItem.Name = "OpenFileToolStripMenuItem";
            this.OpenFileToolStripMenuItem.Size = new System.Drawing.Size(281, 44);
            this.OpenFileToolStripMenuItem.Text = "Відкрити";
            this.OpenFileToolStripMenuItem.Click += new System.EventHandler(this.OpenFileToolStripMenuItem_Click);
            // 
            // SaveFileToolStripMenuItem
            // 
            this.SaveFileToolStripMenuItem.Name = "SaveFileToolStripMenuItem";
            this.SaveFileToolStripMenuItem.Size = new System.Drawing.Size(281, 44);
            this.SaveFileToolStripMenuItem.Text = "Зберегти";
            this.SaveFileToolStripMenuItem.Click += new System.EventHandler(this.SaveFileToolStripMenuItem_Click);
            // 
            // SaveFileAsToolStripMenuItem
            // 
            this.SaveFileAsToolStripMenuItem.Name = "SaveFileAsToolStripMenuItem";
            this.SaveFileAsToolStripMenuItem.Size = new System.Drawing.Size(281, 44);
            this.SaveFileAsToolStripMenuItem.Text = "Зберегти як";
            this.SaveFileAsToolStripMenuItem.Click += new System.EventHandler(this.SaveFileAsToolStripMenuItem_Click);
            // 
            // EditToolStripMenuItem
            // 
            this.EditToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddToolStripMenuItem,
            this.DeleteToolStripMenuItem,
            this.clearTableToolStripMenuItem});
            this.EditToolStripMenuItem.Name = "EditToolStripMenuItem";
            this.EditToolStripMenuItem.Size = new System.Drawing.Size(171, 36);
            this.EditToolStripMenuItem.Text = "Редагування";
            // 
            // AddToolStripMenuItem
            // 
            this.AddToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddRowToolStripMenuItem,
            this.AddColumToolStripMenuItem});
            this.AddToolStripMenuItem.Name = "AddToolStripMenuItem";
            this.AddToolStripMenuItem.Size = new System.Drawing.Size(356, 44);
            this.AddToolStripMenuItem.Text = "Додати";
            // 
            // AddRowToolStripMenuItem
            // 
            this.AddRowToolStripMenuItem.Name = "AddRowToolStripMenuItem";
            this.AddRowToolStripMenuItem.Size = new System.Drawing.Size(253, 44);
            this.AddRowToolStripMenuItem.Text = "Рядок";
            this.AddRowToolStripMenuItem.Click += new System.EventHandler(this.AddRowToolStripMenuItem_Click);
            // 
            // AddColumToolStripMenuItem
            // 
            this.AddColumToolStripMenuItem.Name = "AddColumToolStripMenuItem";
            this.AddColumToolStripMenuItem.Size = new System.Drawing.Size(253, 44);
            this.AddColumToolStripMenuItem.Text = "Стовпчик";
            this.AddColumToolStripMenuItem.Click += new System.EventHandler(this.AddColumToolStripMenuItem_Click);
            // 
            // DeleteToolStripMenuItem
            // 
            this.DeleteToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DeleteRowToolStripMenuItem1,
            this.DeleteColToolStripMenuItem1});
            this.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem";
            this.DeleteToolStripMenuItem.Size = new System.Drawing.Size(356, 44);
            this.DeleteToolStripMenuItem.Text = "Видалити";
            // 
            // DeleteRowToolStripMenuItem1
            // 
            this.DeleteRowToolStripMenuItem1.Name = "DeleteRowToolStripMenuItem1";
            this.DeleteRowToolStripMenuItem1.Size = new System.Drawing.Size(253, 44);
            this.DeleteRowToolStripMenuItem1.Text = "Рядок";
            this.DeleteRowToolStripMenuItem1.Click += new System.EventHandler(this.DeleteRowToolStripMenuItem1_Click);
            // 
            // DeleteColToolStripMenuItem1
            // 
            this.DeleteColToolStripMenuItem1.Name = "DeleteColToolStripMenuItem1";
            this.DeleteColToolStripMenuItem1.Size = new System.Drawing.Size(253, 44);
            this.DeleteColToolStripMenuItem1.Text = "Стовпчик";
            this.DeleteColToolStripMenuItem1.Click += new System.EventHandler(this.DeleteColToolStripMenuItem1_Click);
            // 
            // clearTableToolStripMenuItem
            // 
            this.clearTableToolStripMenuItem.Name = "clearTableToolStripMenuItem";
            this.clearTableToolStripMenuItem.Size = new System.Drawing.Size(356, 44);
            this.clearTableToolStripMenuItem.Text = "Очистити таблицю";
            this.clearTableToolStripMenuItem.Click += new System.EventHandler(this.clearTableToolStripMenuItem_Click);
            // 
            // informationToolStripMenuItem
            // 
            this.informationToolStripMenuItem.Name = "informationToolStripMenuItem";
            this.informationToolStripMenuItem.Size = new System.Drawing.Size(121, 36);
            this.informationToolStripMenuItem.Text = "Довідка";
            this.informationToolStripMenuItem.Click += new System.EventHandler(this.informationToolStripMenuItem_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dataGridView1.Location = new System.Drawing.Point(4, 2);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 82;
            this.dataGridView1.RowTemplate.Height = 41;
            this.dataGridView1.Size = new System.Drawing.Size(883, 483);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView1_CellBeginEdit);
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEndEdit);
            this.dataGridView1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter);
            this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.textBox1.Location = new System.Drawing.Point(274, 2);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(613, 54);
            this.textBox1.TabIndex = 2;
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Segoe Script", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(164, 0);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 66);
            this.label1.TabIndex = 3;
            this.label1.Text = "f(x)";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBox1, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 40);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 66F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(891, 66);
            this.tableLayoutPanel1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Microsoft New Tai Lue", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(4, 0);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 66);
            this.label2.TabIndex = 4;
            this.label2.Text = "A1";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.dataGridView1, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 106);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(891, 487);
            this.tableLayoutPanel2.TabIndex = 5;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xml";
            this.saveFileDialog1.Filter = "XML Files (*.xml)|*.xml";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // MyExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 32F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(891, 593);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "MyExcel";
            this.Text = "MyExcel";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MyExcel_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem FileToolStripMenuItem;
        private ToolStripMenuItem OpenFileToolStripMenuItem;
        private ToolStripMenuItem SaveFileToolStripMenuItem;
        private ToolStripMenuItem EditToolStripMenuItem;
        private ToolStripMenuItem AddToolStripMenuItem;
        private ToolStripMenuItem AddRowToolStripMenuItem;
        private ToolStripMenuItem AddColumToolStripMenuItem;
        private ToolStripMenuItem DeleteToolStripMenuItem;
        private ToolStripMenuItem DeleteRowToolStripMenuItem1;
        private ToolStripMenuItem DeleteColToolStripMenuItem1;
        private DataGridView dataGridView1;
        private TextBox textBox1;
        private Label label1;
        private TableLayoutPanel tableLayoutPanel1;
        private TableLayoutPanel tableLayoutPanel2;
        private Label label2;
        private ToolStripMenuItem SaveFileAsToolStripMenuItem;
        private SaveFileDialog saveFileDialog1;
        private OpenFileDialog openFileDialog1;
        private ToolStripMenuItem clearTableToolStripMenuItem;
        private ToolStripMenuItem informationToolStripMenuItem;
        private ToolStripMenuItem newFileToolStripMenuItem;
    }
}