using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CheckTheWordDocument
{
    public partial class Form2 : Form
    {
        IniFile INI = new IniFile("config.ini");
        List<Style> styles_from_template = new List<Style>(); //all styles
        //List<Style> selected_styles = new List<Style>();
        List<string> selected_style_names = new List<string>(); //styles checked by user

        public Form2(string pathtotemplate)
        {
            InitializeComponent();
            styles_from_template = Checker.GetStylesPart(pathtotemplate, true, false);
            foreach (Style s in styles_from_template)
                checkedListBox1.Items.Add(s.StyleName.Val);
            auto_read();
        }

        private void auto_read()
        {
            /*if (INI.KeyExistsINI("SettingForm1", "Width"))
                numericUpDown2.Value = int.Parse(INI.ReadINI("SettingForm1", "Height"));
            else
                numericUpDown1.Value = this.MinimumSize.Height;

            if (INI.KeyExistsINI("SettingForm1", "Height"))
                numericUpDown1.Value = int.Parse(INI.ReadINI("SettingForm1", "Width"));
            else
                numericUpDown2.Value = this.MinimumSize.Width;

            if (INI.KeyExistsINI("SettingForm1", "Width"))
                textBox1.Text = INI.ReadINI("Other", "Text");

            this.Height = int.Parse(numericUpDown1.Value.ToString());
            this.Width = int.Parse(numericUpDown2.Value.ToString());*/


            /*INI.WriteINI("SettingForm1", "Height", numericUpDown2.Value.ToString());
            INI.WriteINI("SettingForm1", "Width", numericUpDown1.Value.ToString());
            INI.WriteINI("Other", "Text", textBox1.Text);*/


            /*INI.WriteINI("SettingForm1", "Height", numericUpDown2.Value.ToString());
            INI.WriteINI("SettingForm1", "Width", numericUpDown1.Value.ToString());
            this.Height = int.Parse(numericUpDown1.Value.ToString());
            this.Width = int.Parse(numericUpDown2.Value.ToString());*/
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (int index in checkedListBox1.CheckedIndices)
            {
                selected_style_names.Add(styles_from_template[index].StyleName.Val);
            }
            Checker.AddStyleNames(selected_style_names);
            this.Close();
        }
    }
}
