﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddValue
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Console.WriteLine("button clicked\n");
			if (this.openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
			this.excelWrapper1.file_open(this.openFileDialog1.FileName);
			Console.WriteLine("END button clicked\n");
		}

		private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
		{

		}
	}
}