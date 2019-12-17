namespace ExcelGridDataProviderEPPlus
{
    partial class SettingsDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.fileLabel = new System.Windows.Forms.Label();
            this.fileNameTextBox = new System.Windows.Forms.TextBox();
            this.excelGridDataSettingsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.fileBrowseButton = new System.Windows.Forms.Button();
            this.settingsGroup = new System.Windows.Forms.GroupBox();
            this.specificRangeWorksheetComboBox = new System.Windows.Forms.ComboBox();
            this.worksheetsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.worksheetNameComboBox = new System.Windows.Forms.ComboBox();
            this.namedRangeComboBox = new System.Windows.Forms.ComboBox();
            this.namedRangesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.specificRangeTextBox = new System.Windows.Forms.TextBox();
            this.namedRangeRadioButton = new System.Windows.Forms.RadioButton();
            this.specificRangeRadioButton = new System.Windows.Forms.RadioButton();
            this.wholeWorksheetRadioButton = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.excelGridDataSettingsBindingSource)).BeginInit();
            this.settingsGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.worksheetsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.namedRangesBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(412, 182);
            this.okButton.Margin = new System.Windows.Forms.Padding(4);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(100, 28);
            this.okButton.TabIndex = 4;
            this.okButton.Text = "&OK";
            this.okButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(520, 182);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 28);
            this.cancelButton.TabIndex = 5;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // fileLabel
            // 
            this.fileLabel.AutoSize = true;
            this.fileLabel.Location = new System.Drawing.Point(16, 11);
            this.fileLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.fileLabel.Name = "fileLabel";
            this.fileLabel.Size = new System.Drawing.Size(75, 17);
            this.fileLabel.TabIndex = 0;
            this.fileLabel.Text = "&File Name:";
            // 
            // fileNameTextBox
            // 
            this.fileNameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fileNameTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.excelGridDataSettingsBindingSource, "FileName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.fileNameTextBox.Location = new System.Drawing.Point(100, 7);
            this.fileNameTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.fileNameTextBox.Name = "fileNameTextBox";
            this.fileNameTextBox.Size = new System.Drawing.Size(464, 22);
            this.fileNameTextBox.TabIndex = 1;
            // 
            // excelGridDataSettingsBindingSource
            // 
            this.excelGridDataSettingsBindingSource.DataSource = typeof(ExcelGridDataSettings);
            // 
            // fileBrowseButton
            // 
            this.fileBrowseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.fileBrowseButton.Location = new System.Drawing.Point(573, 7);
            this.fileBrowseButton.Margin = new System.Windows.Forms.Padding(4);
            this.fileBrowseButton.Name = "fileBrowseButton";
            this.fileBrowseButton.Size = new System.Drawing.Size(45, 26);
            this.fileBrowseButton.TabIndex = 2;
            this.fileBrowseButton.Text = "...";
            this.fileBrowseButton.UseVisualStyleBackColor = true;
            this.fileBrowseButton.Click += new System.EventHandler(this.fileBrowseButton_Click);
            // 
            // settingsGroup
            // 
            this.settingsGroup.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.settingsGroup.Controls.Add(this.specificRangeWorksheetComboBox);
            this.settingsGroup.Controls.Add(this.worksheetNameComboBox);
            this.settingsGroup.Controls.Add(this.namedRangeComboBox);
            this.settingsGroup.Controls.Add(this.specificRangeTextBox);
            this.settingsGroup.Controls.Add(this.namedRangeRadioButton);
            this.settingsGroup.Controls.Add(this.specificRangeRadioButton);
            this.settingsGroup.Controls.Add(this.wholeWorksheetRadioButton);
            this.settingsGroup.Location = new System.Drawing.Point(20, 41);
            this.settingsGroup.Margin = new System.Windows.Forms.Padding(4);
            this.settingsGroup.Name = "settingsGroup";
            this.settingsGroup.Padding = new System.Windows.Forms.Padding(4);
            this.settingsGroup.Size = new System.Drawing.Size(599, 118);
            this.settingsGroup.TabIndex = 3;
            this.settingsGroup.TabStop = false;
            this.settingsGroup.Text = "&Use data in";
            // 
            // specificRangeWorksheetComboBox
            // 
            this.specificRangeWorksheetComboBox.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.excelGridDataSettingsBindingSource, "Worksheet", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.specificRangeWorksheetComboBox.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.excelGridDataSettingsBindingSource, "IsSpecificRange", true));
            this.specificRangeWorksheetComboBox.DataSource = this.worksheetsBindingSource;
            this.specificRangeWorksheetComboBox.FormattingEnabled = true;
            this.specificRangeWorksheetComboBox.Location = new System.Drawing.Point(168, 79);
            this.specificRangeWorksheetComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.specificRangeWorksheetComboBox.Name = "specificRangeWorksheetComboBox";
            this.specificRangeWorksheetComboBox.Size = new System.Drawing.Size(197, 24);
            this.specificRangeWorksheetComboBox.TabIndex = 5;
            // 
            // worksheetsBindingSource
            // 
            this.worksheetsBindingSource.DataMember = "Worksheets";
            this.worksheetsBindingSource.DataSource = this.excelGridDataSettingsBindingSource;
            // 
            // worksheetNameComboBox
            // 
            this.worksheetNameComboBox.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.excelGridDataSettingsBindingSource, "Worksheet", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.worksheetNameComboBox.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.excelGridDataSettingsBindingSource, "IsWorksheetRange", true));
            this.worksheetNameComboBox.DataSource = this.worksheetsBindingSource;
            this.worksheetNameComboBox.FormattingEnabled = true;
            this.worksheetNameComboBox.Location = new System.Drawing.Point(168, 22);
            this.worksheetNameComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.worksheetNameComboBox.Name = "worksheetNameComboBox";
            this.worksheetNameComboBox.Size = new System.Drawing.Size(197, 24);
            this.worksheetNameComboBox.TabIndex = 1;
            // 
            // namedRangeComboBox
            // 
            this.namedRangeComboBox.DataBindings.Add(new System.Windows.Forms.Binding("SelectedItem", this.excelGridDataSettingsBindingSource, "NamedRange", true));
            this.namedRangeComboBox.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.excelGridDataSettingsBindingSource, "IsNamedRange", true));
            this.namedRangeComboBox.DataSource = this.namedRangesBindingSource;
            this.namedRangeComboBox.FormattingEnabled = true;
            this.namedRangeComboBox.Location = new System.Drawing.Point(168, 50);
            this.namedRangeComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.namedRangeComboBox.Name = "namedRangeComboBox";
            this.namedRangeComboBox.Size = new System.Drawing.Size(197, 24);
            this.namedRangeComboBox.TabIndex = 3;
            // 
            // namedRangesBindingSource
            // 
            this.namedRangesBindingSource.DataMember = "NamedRanges";
            this.namedRangesBindingSource.DataSource = this.excelGridDataSettingsBindingSource;
            // 
            // specificRangeTextBox
            // 
            this.specificRangeTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.specificRangeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.excelGridDataSettingsBindingSource, "IsSpecificRange", true));
            this.specificRangeTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.excelGridDataSettingsBindingSource, "SpecificRange", true));
            this.specificRangeTextBox.Location = new System.Drawing.Point(375, 80);
            this.specificRangeTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.specificRangeTextBox.Name = "specificRangeTextBox";
            this.specificRangeTextBox.Size = new System.Drawing.Size(215, 22);
            this.specificRangeTextBox.TabIndex = 6;
            // 
            // namedRangeRadioButton
            // 
            this.namedRangeRadioButton.AutoCheck = false;
            this.namedRangeRadioButton.AutoSize = true;
            this.namedRangeRadioButton.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.excelGridDataSettingsBindingSource, "IsNamedRange", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.namedRangeRadioButton.Location = new System.Drawing.Point(8, 52);
            this.namedRangeRadioButton.Margin = new System.Windows.Forms.Padding(4);
            this.namedRangeRadioButton.Name = "namedRangeRadioButton";
            this.namedRangeRadioButton.Size = new System.Drawing.Size(124, 21);
            this.namedRangeRadioButton.TabIndex = 2;
            this.namedRangeRadioButton.Text = "&Named Range:";
            this.namedRangeRadioButton.UseVisualStyleBackColor = true;
            this.namedRangeRadioButton.Click += new System.EventHandler(this.namedRangeRadioButton_Click);
            // 
            // specificRangeRadioButton
            // 
            this.specificRangeRadioButton.AutoCheck = false;
            this.specificRangeRadioButton.AutoSize = true;
            this.specificRangeRadioButton.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.excelGridDataSettingsBindingSource, "IsSpecificRange", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.specificRangeRadioButton.Location = new System.Drawing.Point(8, 80);
            this.specificRangeRadioButton.Margin = new System.Windows.Forms.Padding(4);
            this.specificRangeRadioButton.Name = "specificRangeRadioButton";
            this.specificRangeRadioButton.Size = new System.Drawing.Size(128, 21);
            this.specificRangeRadioButton.TabIndex = 4;
            this.specificRangeRadioButton.Text = "&Specific Range:";
            this.specificRangeRadioButton.UseVisualStyleBackColor = true;
            this.specificRangeRadioButton.Click += new System.EventHandler(this.specificRangeRadioButton_Click);
            // 
            // wholeWorksheetRadioButton
            // 
            this.wholeWorksheetRadioButton.AutoCheck = false;
            this.wholeWorksheetRadioButton.AutoSize = true;
            this.wholeWorksheetRadioButton.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.excelGridDataSettingsBindingSource, "IsWorksheetRange", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.wholeWorksheetRadioButton.Location = new System.Drawing.Point(8, 23);
            this.wholeWorksheetRadioButton.Margin = new System.Windows.Forms.Padding(4);
            this.wholeWorksheetRadioButton.Name = "wholeWorksheetRadioButton";
            this.wholeWorksheetRadioButton.Size = new System.Drawing.Size(145, 21);
            this.wholeWorksheetRadioButton.TabIndex = 0;
            this.wholeWorksheetRadioButton.Text = "W&hole Worksheet:";
            this.wholeWorksheetRadioButton.UseVisualStyleBackColor = true;
            this.wholeWorksheetRadioButton.Click += new System.EventHandler(this.wholeWorksheetRadioButton_Click);
            // 
            // SettingsDialog
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(635, 225);
            this.Controls.Add(this.settingsGroup);
            this.Controls.Add(this.fileBrowseButton);
            this.Controls.Add(this.fileNameTextBox);
            this.Controls.Add(this.fileLabel);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(634, 248);
            this.Name = "SettingsDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Excel Settings";
            this.Load += new System.EventHandler(this.SettingsDialog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.excelGridDataSettingsBindingSource)).EndInit();
            this.settingsGroup.ResumeLayout(false);
            this.settingsGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.worksheetsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.namedRangesBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label fileLabel;
        private System.Windows.Forms.TextBox fileNameTextBox;
        private System.Windows.Forms.Button fileBrowseButton;
        private System.Windows.Forms.GroupBox settingsGroup;
        private System.Windows.Forms.RadioButton namedRangeRadioButton;
        private System.Windows.Forms.RadioButton specificRangeRadioButton;
        private System.Windows.Forms.RadioButton wholeWorksheetRadioButton;
        private System.Windows.Forms.ComboBox namedRangeComboBox;
        private System.Windows.Forms.TextBox specificRangeTextBox;
        private System.Windows.Forms.ComboBox worksheetNameComboBox;
        private System.Windows.Forms.BindingSource excelGridDataSettingsBindingSource;
        private System.Windows.Forms.BindingSource worksheetsBindingSource;
        private System.Windows.Forms.BindingSource namedRangesBindingSource;
        private System.Windows.Forms.ComboBox specificRangeWorksheetComboBox;
    }
}