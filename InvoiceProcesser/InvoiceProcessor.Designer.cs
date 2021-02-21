namespace InvoiceProcessor
{
    partial class mainWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainWindow));
            this.selectFileButton = new System.Windows.Forms.Button();
            this.createFileButton = new System.Windows.Forms.Button();
            this.invoiceDropPanel = new System.Windows.Forms.Panel();
            this.invoiceDragAreaLabel = new System.Windows.Forms.Label();
            this.fileSelectedLabel = new System.Windows.Forms.Label();
            this.fileUpdatedLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.invoiceDropPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // selectFileButton
            // 
            this.selectFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.selectFileButton.Location = new System.Drawing.Point(52, 34);
            this.selectFileButton.Name = "selectFileButton";
            this.selectFileButton.Size = new System.Drawing.Size(180, 23);
            this.selectFileButton.TabIndex = 0;
            this.selectFileButton.Text = "Select Existing Summary File";
            this.selectFileButton.UseVisualStyleBackColor = true;
            this.selectFileButton.Click += new System.EventHandler(this.selectFileButton_Click);
            // 
            // createFileButton
            // 
            this.createFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.createFileButton.Location = new System.Drawing.Point(52, 104);
            this.createFileButton.Name = "createFileButton";
            this.createFileButton.Size = new System.Drawing.Size(180, 23);
            this.createFileButton.TabIndex = 1;
            this.createFileButton.Text = "Create New Summary File";
            this.createFileButton.UseVisualStyleBackColor = true;
            this.createFileButton.Click += new System.EventHandler(this.createFileButton_Click);
            // 
            // invoiceDropPanel
            // 
            this.invoiceDropPanel.AllowDrop = true;
            this.invoiceDropPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.invoiceDropPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.invoiceDropPanel.Controls.Add(this.invoiceDragAreaLabel);
            this.invoiceDropPanel.Location = new System.Drawing.Point(52, 159);
            this.invoiceDropPanel.Name = "invoiceDropPanel";
            this.invoiceDropPanel.Size = new System.Drawing.Size(180, 170);
            this.invoiceDropPanel.TabIndex = 2;
            this.invoiceDropPanel.DragDrop += new System.Windows.Forms.DragEventHandler(this.invoiceDropPanel_DragDrop);
            this.invoiceDropPanel.DragEnter += new System.Windows.Forms.DragEventHandler(this.invoiceDropPanel_DragEnter);
            // 
            // invoiceDragAreaLabel
            // 
            this.invoiceDragAreaLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.invoiceDragAreaLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.invoiceDragAreaLabel.Location = new System.Drawing.Point(3, 67);
            this.invoiceDragAreaLabel.Name = "invoiceDragAreaLabel";
            this.invoiceDragAreaLabel.Size = new System.Drawing.Size(172, 23);
            this.invoiceDragAreaLabel.TabIndex = 0;
            this.invoiceDragAreaLabel.Text = "Drag Invoices Here..";
            this.invoiceDragAreaLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // fileSelectedLabel
            // 
            this.fileSelectedLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileSelectedLabel.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.fileSelectedLabel.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.fileSelectedLabel.Location = new System.Drawing.Point(52, 60);
            this.fileSelectedLabel.Name = "fileSelectedLabel";
            this.fileSelectedLabel.Size = new System.Drawing.Size(180, 23);
            this.fileSelectedLabel.TabIndex = 3;
            this.fileSelectedLabel.Text = "No File Selected";
            this.fileSelectedLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // fileUpdatedLabel
            // 
            this.fileUpdatedLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.fileUpdatedLabel.AutoSize = true;
            this.fileUpdatedLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileUpdatedLabel.ForeColor = System.Drawing.Color.Green;
            this.fileUpdatedLabel.Location = new System.Drawing.Point(84, 344);
            this.fileUpdatedLabel.Name = "fileUpdatedLabel";
            this.fileUpdatedLabel.Size = new System.Drawing.Size(111, 13);
            this.fileUpdatedLabel.TabIndex = 4;
            this.fileUpdatedLabel.Text = "Update Completed";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(87, 360);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(108, 16);
            this.progressBar.TabIndex = 5;
            // 
            // mainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(287, 388);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.fileUpdatedLabel);
            this.Controls.Add(this.fileSelectedLabel);
            this.Controls.Add(this.invoiceDropPanel);
            this.Controls.Add(this.createFileButton);
            this.Controls.Add(this.selectFileButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "mainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Invoice Processor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.mainWindow_FormClosing_1);
            this.invoiceDropPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button selectFileButton;
        private System.Windows.Forms.Button createFileButton;
        private System.Windows.Forms.Panel invoiceDropPanel;
        private System.Windows.Forms.Label fileSelectedLabel;
        private System.Windows.Forms.Label invoiceDragAreaLabel;
        private System.Windows.Forms.Label fileUpdatedLabel;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

