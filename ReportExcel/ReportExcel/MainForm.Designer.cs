
namespace ReportExcel
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.searchingWords_ListBox = new System.Windows.Forms.ListBox();
            this.addSearchingWord_Button = new System.Windows.Forms.Button();
            this.addData_Button = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.excelFiles_ListBox = new System.Windows.Forms.ListBox();
            this.rowName_Label = new System.Windows.Forms.Label();
            this.columnName_Label = new System.Windows.Forms.Label();
            this.productName_TextBox = new System.Windows.Forms.TextBox();
            this.rowName_TextBox = new System.Windows.Forms.TextBox();
            this.columnName_TextBox = new System.Windows.Forms.TextBox();
            this.productPrice_TextBox = new System.Windows.Forms.TextBox();
            this.rowPrice_Label = new System.Windows.Forms.Label();
            this.columnPrice_Label = new System.Windows.Forms.Label();
            this.rowPrice_TextBox = new System.Windows.Forms.TextBox();
            this.columnPrice_TextBox = new System.Windows.Forms.TextBox();
            this.savePosition_Button = new System.Windows.Forms.Button();
            this.excelFiles_Label = new System.Windows.Forms.Label();
            this.addExcelFile_Button = new System.Windows.Forms.Button();
            this.deleteSearchingWord_Button = new System.Windows.Forms.Button();
            this.searchingWord_TextBox = new System.Windows.Forms.TextBox();
            this.discount_Label = new System.Windows.Forms.Label();
            this.discount_TextBox = new System.Windows.Forms.TextBox();
            this.productDiscount_TextBox = new System.Windows.Forms.TextBox();
            this.exportData_Button = new System.Windows.Forms.Button();
            this.percent_Label = new System.Windows.Forms.Label();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // searchingWords_ListBox
            // 
            this.searchingWords_ListBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(20)))), ((int)(((byte)(39)))));
            this.searchingWords_ListBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.searchingWords_ListBox.ForeColor = System.Drawing.Color.Gainsboro;
            this.searchingWords_ListBox.FormattingEnabled = true;
            this.searchingWords_ListBox.ItemHeight = 21;
            this.searchingWords_ListBox.Location = new System.Drawing.Point(380, 12);
            this.searchingWords_ListBox.Name = "searchingWords_ListBox";
            this.searchingWords_ListBox.Size = new System.Drawing.Size(814, 151);
            this.searchingWords_ListBox.TabIndex = 4;
            // 
            // addSearchingWord_Button
            // 
            this.addSearchingWord_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.addSearchingWord_Button.Location = new System.Drawing.Point(32, 77);
            this.addSearchingWord_Button.Name = "addSearchingWord_Button";
            this.addSearchingWord_Button.Size = new System.Drawing.Size(168, 86);
            this.addSearchingWord_Button.TabIndex = 5;
            this.addSearchingWord_Button.Text = "Добавить искомое слово в список";
            this.addSearchingWord_Button.UseVisualStyleBackColor = true;
            this.addSearchingWord_Button.Click += new System.EventHandler(this.AddSearchingWord_Button_Click);
            // 
            // addData_Button
            // 
            this.addData_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.addData_Button.Location = new System.Drawing.Point(794, 565);
            this.addData_Button.Name = "addData_Button";
            this.addData_Button.Size = new System.Drawing.Size(400, 93);
            this.addData_Button.TabIndex = 6;
            this.addData_Button.Text = "Загрузка данных из excel файла";
            this.addData_Button.UseVisualStyleBackColor = true;
            this.addData_Button.Click += new System.EventHandler(this.AddData_Button_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // excelFiles_ListBox
            // 
            this.excelFiles_ListBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(20)))), ((int)(((byte)(39)))));
            this.excelFiles_ListBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.excelFiles_ListBox.ForeColor = System.Drawing.Color.Gainsboro;
            this.excelFiles_ListBox.FormattingEnabled = true;
            this.excelFiles_ListBox.ItemHeight = 21;
            this.excelFiles_ListBox.Location = new System.Drawing.Point(380, 219);
            this.excelFiles_ListBox.Name = "excelFiles_ListBox";
            this.excelFiles_ListBox.Size = new System.Drawing.Size(814, 340);
            this.excelFiles_ListBox.TabIndex = 7;
            this.excelFiles_ListBox.SelectedIndexChanged += new System.EventHandler(this.ExcelFiles_ListBox_SelectedIndexChanged);
            // 
            // rowName_Label
            // 
            this.rowName_Label.AutoSize = true;
            this.rowName_Label.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rowName_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.rowName_Label.Location = new System.Drawing.Point(79, 267);
            this.rowName_Label.Name = "rowName_Label";
            this.rowName_Label.Size = new System.Drawing.Size(79, 23);
            this.rowName_Label.TabIndex = 8;
            this.rowName_Label.Text = "Строка:";
            // 
            // columnName_Label
            // 
            this.columnName_Label.AutoSize = true;
            this.columnName_Label.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.columnName_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.columnName_Label.Location = new System.Drawing.Point(247, 267);
            this.columnName_Label.Name = "columnName_Label";
            this.columnName_Label.Size = new System.Drawing.Size(91, 23);
            this.columnName_Label.TabIndex = 9;
            this.columnName_Label.Text = "Столбец:";
            // 
            // productName_TextBox
            // 
            this.productName_TextBox.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.productName_TextBox.Location = new System.Drawing.Point(32, 219);
            this.productName_TextBox.Name = "productName_TextBox";
            this.productName_TextBox.ReadOnly = true;
            this.productName_TextBox.Size = new System.Drawing.Size(339, 32);
            this.productName_TextBox.TabIndex = 10;
            this.productName_TextBox.Text = "Наименование продукта";
            this.productName_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // rowName_TextBox
            // 
            this.rowName_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rowName_TextBox.Location = new System.Drawing.Point(83, 293);
            this.rowName_TextBox.Name = "rowName_TextBox";
            this.rowName_TextBox.Size = new System.Drawing.Size(50, 29);
            this.rowName_TextBox.TabIndex = 11;
            // 
            // columnName_TextBox
            // 
            this.columnName_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.columnName_TextBox.Location = new System.Drawing.Point(251, 293);
            this.columnName_TextBox.Name = "columnName_TextBox";
            this.columnName_TextBox.Size = new System.Drawing.Size(50, 29);
            this.columnName_TextBox.TabIndex = 12;
            // 
            // productPrice_TextBox
            // 
            this.productPrice_TextBox.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.productPrice_TextBox.Location = new System.Drawing.Point(32, 334);
            this.productPrice_TextBox.Name = "productPrice_TextBox";
            this.productPrice_TextBox.ReadOnly = true;
            this.productPrice_TextBox.Size = new System.Drawing.Size(339, 32);
            this.productPrice_TextBox.TabIndex = 13;
            this.productPrice_TextBox.Text = "Цена Продукта";
            this.productPrice_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // rowPrice_Label
            // 
            this.rowPrice_Label.AutoSize = true;
            this.rowPrice_Label.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rowPrice_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.rowPrice_Label.Location = new System.Drawing.Point(79, 377);
            this.rowPrice_Label.Name = "rowPrice_Label";
            this.rowPrice_Label.Size = new System.Drawing.Size(79, 23);
            this.rowPrice_Label.TabIndex = 14;
            this.rowPrice_Label.Text = "Строка:";
            // 
            // columnPrice_Label
            // 
            this.columnPrice_Label.AutoSize = true;
            this.columnPrice_Label.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.columnPrice_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.columnPrice_Label.Location = new System.Drawing.Point(247, 377);
            this.columnPrice_Label.Name = "columnPrice_Label";
            this.columnPrice_Label.Size = new System.Drawing.Size(91, 23);
            this.columnPrice_Label.TabIndex = 15;
            this.columnPrice_Label.Text = "Столбец:";
            // 
            // rowPrice_TextBox
            // 
            this.rowPrice_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rowPrice_TextBox.Location = new System.Drawing.Point(83, 403);
            this.rowPrice_TextBox.Name = "rowPrice_TextBox";
            this.rowPrice_TextBox.Size = new System.Drawing.Size(50, 29);
            this.rowPrice_TextBox.TabIndex = 16;
            // 
            // columnPrice_TextBox
            // 
            this.columnPrice_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.columnPrice_TextBox.Location = new System.Drawing.Point(251, 403);
            this.columnPrice_TextBox.Name = "columnPrice_TextBox";
            this.columnPrice_TextBox.Size = new System.Drawing.Size(50, 29);
            this.columnPrice_TextBox.TabIndex = 17;
            // 
            // savePosition_Button
            // 
            this.savePosition_Button.BackColor = System.Drawing.Color.White;
            this.savePosition_Button.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.savePosition_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.savePosition_Button.Location = new System.Drawing.Point(32, 520);
            this.savePosition_Button.Name = "savePosition_Button";
            this.savePosition_Button.Size = new System.Drawing.Size(342, 93);
            this.savePosition_Button.TabIndex = 18;
            this.savePosition_Button.Text = "Сохранить значения строки и столбца для начало чтения excel файла";
            this.savePosition_Button.UseVisualStyleBackColor = true;
            this.savePosition_Button.Click += new System.EventHandler(this.SavePosition_Button_Click);
            // 
            // excelFiles_Label
            // 
            this.excelFiles_Label.AutoSize = true;
            this.excelFiles_Label.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.excelFiles_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.excelFiles_Label.Location = new System.Drawing.Point(376, 195);
            this.excelFiles_Label.Name = "excelFiles_Label";
            this.excelFiles_Label.Size = new System.Drawing.Size(295, 21);
            this.excelFiles_Label.TabIndex = 19;
            this.excelFiles_Label.Text = "Файлы подготовленные для поиска";
            // 
            // addExcelFile_Button
            // 
            this.addExcelFile_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.addExcelFile_Button.Location = new System.Drawing.Point(380, 565);
            this.addExcelFile_Button.Name = "addExcelFile_Button";
            this.addExcelFile_Button.Size = new System.Drawing.Size(400, 93);
            this.addExcelFile_Button.TabIndex = 20;
            this.addExcelFile_Button.Text = "Добавление Excel файла в список";
            this.addExcelFile_Button.UseVisualStyleBackColor = true;
            this.addExcelFile_Button.Click += new System.EventHandler(this.AddExcelFile_Button_Click);
            // 
            // deleteSearchingWord_Button
            // 
            this.deleteSearchingWord_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.deleteSearchingWord_Button.Location = new System.Drawing.Point(206, 77);
            this.deleteSearchingWord_Button.Name = "deleteSearchingWord_Button";
            this.deleteSearchingWord_Button.Size = new System.Drawing.Size(168, 86);
            this.deleteSearchingWord_Button.TabIndex = 21;
            this.deleteSearchingWord_Button.Text = "Удалить искомое слово";
            this.deleteSearchingWord_Button.UseVisualStyleBackColor = true;
            this.deleteSearchingWord_Button.Click += new System.EventHandler(this.DeleteSearchingWord_Button_Click);
            // 
            // searchingWord_TextBox
            // 
            this.searchingWord_TextBox.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.searchingWord_TextBox.Location = new System.Drawing.Point(32, 12);
            this.searchingWord_TextBox.Name = "searchingWord_TextBox";
            this.searchingWord_TextBox.Size = new System.Drawing.Size(342, 32);
            this.searchingWord_TextBox.TabIndex = 22;
            // 
            // discount_Label
            // 
            this.discount_Label.AutoSize = true;
            this.discount_Label.Font = new System.Drawing.Font("Times New Roman", 14.25F);
            this.discount_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.discount_Label.Location = new System.Drawing.Point(85, 488);
            this.discount_Label.Name = "discount_Label";
            this.discount_Label.Size = new System.Drawing.Size(73, 21);
            this.discount_Label.TabIndex = 23;
            this.discount_Label.Text = "Скидка:";
            // 
            // discount_TextBox
            // 
            this.discount_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F);
            this.discount_TextBox.Location = new System.Drawing.Point(163, 485);
            this.discount_TextBox.Name = "discount_TextBox";
            this.discount_TextBox.Size = new System.Drawing.Size(100, 29);
            this.discount_TextBox.TabIndex = 24;
            // 
            // productDiscount_TextBox
            // 
            this.productDiscount_TextBox.BackColor = System.Drawing.SystemColors.Control;
            this.productDiscount_TextBox.Font = new System.Drawing.Font("Times New Roman", 14.25F);
            this.productDiscount_TextBox.Location = new System.Drawing.Point(32, 438);
            this.productDiscount_TextBox.Name = "productDiscount_TextBox";
            this.productDiscount_TextBox.Size = new System.Drawing.Size(342, 29);
            this.productDiscount_TextBox.TabIndex = 25;
            this.productDiscount_TextBox.Text = "Скидка от поставщика";
            this.productDiscount_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // exportData_Button
            // 
            this.exportData_Button.Font = new System.Drawing.Font("Times New Roman", 14.25F);
            this.exportData_Button.Location = new System.Drawing.Point(794, 664);
            this.exportData_Button.Name = "exportData_Button";
            this.exportData_Button.Size = new System.Drawing.Size(400, 93);
            this.exportData_Button.TabIndex = 26;
            this.exportData_Button.Text = "Выгрузка Excel файла";
            this.exportData_Button.UseVisualStyleBackColor = true;
            this.exportData_Button.Click += new System.EventHandler(this.ExportData_Button_Click);
            // 
            // percent_Label
            // 
            this.percent_Label.AutoSize = true;
            this.percent_Label.Font = new System.Drawing.Font("Times New Roman", 14.25F);
            this.percent_Label.ForeColor = System.Drawing.Color.Gainsboro;
            this.percent_Label.Location = new System.Drawing.Point(269, 488);
            this.percent_Label.Name = "percent_Label";
            this.percent_Label.Size = new System.Drawing.Size(26, 21);
            this.percent_Label.TabIndex = 27;
            this.percent_Label.Text = "%";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(39)))), ((int)(((byte)(58)))));
            this.ClientSize = new System.Drawing.Size(1212, 766);
            this.Controls.Add(this.percent_Label);
            this.Controls.Add(this.exportData_Button);
            this.Controls.Add(this.productDiscount_TextBox);
            this.Controls.Add(this.discount_TextBox);
            this.Controls.Add(this.discount_Label);
            this.Controls.Add(this.searchingWord_TextBox);
            this.Controls.Add(this.deleteSearchingWord_Button);
            this.Controls.Add(this.addExcelFile_Button);
            this.Controls.Add(this.excelFiles_Label);
            this.Controls.Add(this.savePosition_Button);
            this.Controls.Add(this.columnPrice_TextBox);
            this.Controls.Add(this.rowPrice_TextBox);
            this.Controls.Add(this.columnPrice_Label);
            this.Controls.Add(this.rowPrice_Label);
            this.Controls.Add(this.productPrice_TextBox);
            this.Controls.Add(this.columnName_TextBox);
            this.Controls.Add(this.rowName_TextBox);
            this.Controls.Add(this.productName_TextBox);
            this.Controls.Add(this.columnName_Label);
            this.Controls.Add(this.rowName_Label);
            this.Controls.Add(this.addData_Button);
            this.Controls.Add(this.addSearchingWord_Button);
            this.Controls.Add(this.searchingWords_ListBox);
            this.Controls.Add(this.excelFiles_ListBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListBox searchingWords_ListBox;
        private System.Windows.Forms.Button addSearchingWord_Button;
        private System.Windows.Forms.Button addData_Button;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ListBox excelFiles_ListBox;
        private System.Windows.Forms.Label rowName_Label;
        private System.Windows.Forms.Label columnName_Label;
        private System.Windows.Forms.TextBox productName_TextBox;
        private System.Windows.Forms.TextBox rowName_TextBox;
        private System.Windows.Forms.TextBox columnName_TextBox;
        private System.Windows.Forms.TextBox productPrice_TextBox;
        private System.Windows.Forms.Label rowPrice_Label;
        private System.Windows.Forms.Label columnPrice_Label;
        private System.Windows.Forms.TextBox rowPrice_TextBox;
        private System.Windows.Forms.TextBox columnPrice_TextBox;
        private System.Windows.Forms.Button savePosition_Button;
        private System.Windows.Forms.Label excelFiles_Label;
        private System.Windows.Forms.Button addExcelFile_Button;
        private System.Windows.Forms.Button deleteSearchingWord_Button;
        private System.Windows.Forms.TextBox searchingWord_TextBox;
        private System.Windows.Forms.Label discount_Label;
        private System.Windows.Forms.TextBox discount_TextBox;
        private System.Windows.Forms.TextBox productDiscount_TextBox;
        private System.Windows.Forms.Button exportData_Button;
        private System.Windows.Forms.Label percent_Label;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }
}

