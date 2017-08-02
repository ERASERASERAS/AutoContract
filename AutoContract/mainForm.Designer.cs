namespace AutoContract
{
    partial class mainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(mainForm));
            this.contrTypeComboBox = new System.Windows.Forms.ComboBox();
            this.nextButton = new System.Windows.Forms.Button();
            this.exitButtonF = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // contrTypeComboBox
            // 
            resources.ApplyResources(this.contrTypeComboBox, "contrTypeComboBox");
            this.contrTypeComboBox.FormattingEnabled = true;
            this.contrTypeComboBox.Name = "contrTypeComboBox";
            // 
            // nextButton
            // 
            resources.ApplyResources(this.nextButton, "nextButton");
            this.nextButton.Name = "nextButton";
            this.nextButton.UseVisualStyleBackColor = true;
            this.nextButton.Click += new System.EventHandler(this.nextButton_Click);
            // 
            // exitButtonF
            // 
            resources.ApplyResources(this.exitButtonF, "exitButtonF");
            this.exitButtonF.Name = "exitButtonF";
            this.exitButtonF.UseVisualStyleBackColor = true;
            this.exitButtonF.Click += new System.EventHandler(this.exitButtonF_Click);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.exitButtonF);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.contrTypeComboBox);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox contrTypeComboBox;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Button exitButtonF;
    }
}

