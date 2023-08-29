using ExcelMonitor.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMonitor
{
    public partial class Nastavenie : Form
    {
        public Nastavenie()
        {
            InitializeComponent();
            textBox1.Text = Settings.Default.adresa;
            textBox2.Text = Settings.Default.zalozka;
            textBox3.Text = Settings.Default.bunky;
            textBox4.Text = Settings.Default.frekvencia;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.adresa = textBox1.Text;
            Settings.Default.Save();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.zalozka = textBox2.Text;
            Settings.Default.Save();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.bunky = textBox3.Text;
            Settings.Default.Save();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int.Parse(textBox4.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("Zadaná hodnota nie je číslovka!");
                textBox4.Text = "60000";
                textBox4.SelectAll();
                return;
            }
            Settings.Default.frekvencia = textBox4.Text;
            Settings.Default.Save();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Settings.Default.adresa == "")
            {
                MessageBox.Show("Vložte adresu súboru");
                textBox1.Focus();
                return;
            }
            else if (Settings.Default.zalozka == "")
            {
                MessageBox.Show("Vložte názov záložky");
                textBox2.Focus();
                return;
            }
            else if (Settings.Default.bunky == "")
            {
                MessageBox.Show("Vložte zobrazované bunky");
                textBox3.Focus();
                return;
            }
            else if (Settings.Default.frekvencia == "")
            {
                MessageBox.Show("Vložte frekvenciu obnovy");
                textBox4.Focus();
                return;
            }
            else
            {
                Form1 nextForm = new Form1();
                Hide();
                nextForm.ShowDialog();
                Close();
            }
        }
    }
}
