using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace BuildPDF.csproj {
  public partial class Message : Form {
    public Message() {
      InitializeComponent();
    }

    public void Append(string str) {
      rtbMessage.Text += str;
    }

    public void AppendLine(string str) {
      Append(str + "\n");
    }

    private void Message_Load(object sender, EventArgs e) {
      Location = Properties.Settings.Default.MessageLocation;
      Size = Properties.Settings.Default.MessageSize;
    }

    private void Message_FormClosing(object sender, FormClosingEventArgs e) {
      Properties.Settings.Default.MessageLocation = Location;
      Properties.Settings.Default.MessageSize = Size;
      Properties.Settings.Default.Save();
    }
  }
}