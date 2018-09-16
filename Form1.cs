using System;
using System.Collections;
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
		private ArrayList Sheet_list = new ArrayList();

		private void button1_Click(object sender, EventArgs e)
		{
			Console.WriteLine("START Open File\n");

			if (this.openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
				//filename
				//sheet name
				//drop down list
				this.Wrapper1.File_open(this.openFileDialog1.FileName, this.dataGridView1, "",this.comboBox1);
			}
			Console.WriteLine("END Open File\n");
			return;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			Console.WriteLine("START Whats Best! \n");
			try {
				if (this.openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
				this.Wrapper1.File_open_web(this.openFileDialog1.FileName);
			}

			catch {
				Console.WriteLine("Whats best button failed");
			}
			Console.WriteLine("END Whats Best! \n");
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			//this.comboBox1.Items.Add(tmp);
			var selected = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
			this.Wrapper1.File_open(this.openFileDialog1.FileName, this.dataGridView1, selected, this.comboBox1);
		}
	}
}