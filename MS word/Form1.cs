using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MS_word
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            foreach(FontFamily family in FontFamily.Families)
            {
                comp_font.Items.Add(family.Name);
            }
            if (comp_font.Items.Count > 0)
            {
                comp_font.SelectedIndex = 12;
            }
            comp_font.SelectedIndexChanged += comp_font_SelectedIndexChanged;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comb_fontsize.Width = 80;
            comb_fontsize.Height = 25;
        }

        private void comb_fontsize_SelectedIndexChanged(object sender, EventArgs e)
        {
            int size1 = Convert.ToInt32(comb_fontsize.Text);
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily,size1);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            FontStyle style = richTextBox1.SelectionFont.Style;
            if (richTextBox1.SelectionFont.Bold)
            {
                style &= ~FontStyle.Bold;
                toolStripButton1.BackColor = Color.White;

            }
            else
            {
                style |= FontStyle.Bold;
                toolStripButton1.BackColor = Color.FromArgb(173, 216, 230);


            }
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style);

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            FontStyle style = richTextBox1.SelectionFont.Style;
            if (richTextBox1.SelectionFont.Italic)
            {
                style &= ~FontStyle.Italic;

                toolStripButton2.BackColor = Color.White;
            }
            else
            {
                style |= FontStyle.Italic;
                toolStripButton2.BackColor = Color.FromArgb(173, 216, 230);
               
            }
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style);

        }

        private void btn_underline_Click(object sender, EventArgs e)
        {
            FontStyle style = richTextBox1.SelectionFont.Style;
            if (richTextBox1.SelectionFont.Underline)
            {
                style &= ~FontStyle.Underline;

             btn_underline.BackColor = Color.White;
            }
            else
            {
                style |= FontStyle.Underline;
               btn_underline.BackColor = Color.FromArgb(173, 216, 230);

            }
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style);
        }

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {
           ColorDialog cdialog = new ColorDialog();
            DialogResult dr = cdialog.ShowDialog();

            if (dr == DialogResult.OK) 
            {
                richTextBox1.SelectionColor = cdialog.Color;
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionAlignment != null)
            {
                richTextBox1.SelectionAlignment = (HorizontalAlignment.Right);
            }
        }

        private void btn_alinconter_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionAlignment != null)
            {
                richTextBox1.SelectionAlignment = (HorizontalAlignment.Center);
                btn_alinconter.BackColor = Color.FromArgb(173, 216, 230);
            }
            else
            {
                btn_alinconter.BackColor = Color.White;
                richTextBox1.SelectionAlignment = (HorizontalAlignment.Left);
            }
        }

        private void btn_aln_left_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionAlignment != null)
            {
                richTextBox1.SelectionAlignment = (HorizontalAlignment.Left);
                btn_aln_left.BackColor = Color.FromArgb(173, 216, 230);
            }
    
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            int number = Convert.ToInt32(comb_fontsize.Text) + 1;
            comb_fontsize.Text= number.ToString();
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, number);
            
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (comb_fontsize.Text ==  0.ToString())
            {
                MessageBox.Show("لا يمنك تصغير الخط ", "خطا!!",MessageBoxButtons.OK , MessageBoxIcon.Error);
            }
            else
            {
                int number = Convert.ToInt32(comb_fontsize.Text) - 1;
                comb_fontsize.Text = number.ToString();
                richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont.FontFamily, number);
            }

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            ofd.Title = "ادراج صورة";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            ofd.Multiselect = true;
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";

            if (ofd.ShowDialog() == DialogResult.OK) { 
                Image image = Image.FromFile(ofd.FileName); 
                Clipboard.SetImage(image);
                richTextBox1.Paste();
            }
            ofd.ShowDialog();
           
        }

        private void comb_fontsize_Click(object sender, EventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "save as RTF Document";
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath (Environment.SpecialFolder.Desktop);
            saveFileDialog1.Filter = "Rich Text Format|*.rtf";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) { 
                richTextBox1.SaveFile(saveFileDialog1.FileName , RichTextBoxStreamType.RichText);
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ofd.Title = "save as RTF Document";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.Filter = "Rich Text Format|*.rtf";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    richTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.RichText);
                    MessageBox.Show("File loaded suuccessfly");
                }
                catch (Exception ex) { 
                    MessageBox.Show("An error occurred While loading the file " + ex.Message);
                }
            }
        }

        private void comp_font_SelectedIndexChanged(object sender, EventArgs e)
        {
            string font = comp_font.SelectedItem.ToString();
            richTextBox1.SelectionFont = new Font (font , richTextBox1.SelectionFont.Size);
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.Cut();
            }
        }

        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.Copy();
            }
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.Paste();
            }
            richTextBox1.Paste();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            ColorDialog cdialog = new ColorDialog();
            DialogResult dr = cdialog.ShowDialog();

            if (dr == DialogResult.OK)
            {
                richTextBox1.SelectionBackColor = cdialog.Color;
            }
        }

        private bool isarabic(char c)
        {
            return (c >= '\u0600' && c <= '\u06FF') || (c >= '\u0750' && c <= '\u077F');
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox1.Text.Length > 0)
            {
                char firstchar = richTextBox1.Text[0];
                if (isarabic(firstchar))
                {
                    richTextBox1.RightToLeft = RightToLeft.Yes;
                    richTextBox1.SelectionAlignment = HorizontalAlignment.Right;

                }
                else
                {
                    richTextBox1.RightToLeft = RightToLeft.No;
                    richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
                }
            }
        }
    }
}
