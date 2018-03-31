using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CheckTheWordDocument
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /*~Form1()
        {
            Checker.CloseWordDocument();
        }*/

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            

            // Code
            Checker.CloseWordDocument();
            base.OnFormClosing(e);
        } 


        private void buttonOpen_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
            {
                Form2 f2 = new Form2(openFileDialog1.FileName);
                f2.ShowDialog();
                Checker.Start(openFileDialog1.FileName);
                buttonCheck.Enabled = true;
                //buttonSave.Enabled = true;
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {           
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Checker.SaveWordDocument(saveFileDialog1.FileName);
            }
            Checker.CloseWordDocument();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonCheck_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear(); //очищаем листбокс от старых данных
            List<String> l = Checker.CheckAndWrite();
            foreach (String s in l)
            {
                listBox1.Items.Add(s);
            }
            //Checker.CloseWordDocument();
        }
    }
}
