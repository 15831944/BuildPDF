namespace BuildPDF.csproj {
  partial class Message {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing) {
      if (disposing && (components != null)) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
      this.rtbMessage = new System.Windows.Forms.RichTextBox();
      this.SuspendLayout();
      // 
      // rtbMessage
      // 
      this.rtbMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                  | System.Windows.Forms.AnchorStyles.Left)
                  | System.Windows.Forms.AnchorStyles.Right)));
      this.rtbMessage.Location = new System.Drawing.Point(12, 12);
      this.rtbMessage.Name = "rtbMessage";
      this.rtbMessage.Size = new System.Drawing.Size(607, 236);
      this.rtbMessage.TabIndex = 0;
      this.rtbMessage.Text = "";
      // 
      // Message
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(631, 260);
      this.Controls.Add(this.rtbMessage);
      this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.Name = "Message";
      this.Text = "Searching and merging PDFs...";
      this.Load += new System.EventHandler(this.Message_Load);
      this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Message_FormClosing);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.RichTextBox rtbMessage;
  }
}