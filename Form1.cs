using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sorteio
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
      MessageBox.Show(null, "Para sortear, escolhar um arquivo do Excel." + Environment.NewLine + Environment.NewLine +
        "Então clique sobre uma coluna com números inteiros distintos, e então clique em sortear."
        + Environment.NewLine + Environment.NewLine + "Boa sorte!", "Instruções", MessageBoxButtons.OK, MessageBoxIcon.Question);
    }

    internal LoadExcel Le { get; private set; }
    public int ColIndex { get; private set; } = -1;
    public int MinNum { get; private set; }
    public int MaxNum { get; private set; }

    private void button1_Click(object sender, EventArgs e)
    {
      Le = null;
      dataGridView1.DataSource = null;
      ColIndex = -1;
      button2.Enabled = false;
      OpenFileDialog openFileDialog = new OpenFileDialog();
      openFileDialog.Filter = "Excel files (*.xls)|*.xlsx|All files (*.*)|*.*";
      openFileDialog.FilterIndex = 1;
      openFileDialog.RestoreDirectory = true;
      if (openFileDialog.ShowDialog() == DialogResult.OK)
      {
        Le = new LoadExcel(openFileDialog.FileName);
        dataGridView1.DataSource = Le.Tables[0];
      }
    }

    private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
    {
      ColIndex = e.ColumnIndex;
      MinNum = 0;
      MaxNum = 0;
      foreach (DataRow dr in Le.Tables[0].Rows)
      {
        if (int.TryParse(dr[ColIndex].ToString(), out int i))
        {
          if (MinNum > i || MinNum == 0)
            MinNum = i;
          if (MaxNum < i)
            MaxNum = i;
        }
        else
        {
          MinNum = 0;
          MaxNum = 0;
          MessageBox.Show(null, "Atenção, selecione uma coluna com números inteiros distintos!", "Atenção!", MessageBoxButtons.OK);
          break;
        }
      }
      if (MinNum > 0 && MaxNum > 0)
        button2.Enabled = true;
    }

    private void button2_Click(object sender, EventArgs e)
    {
      if (ColIndex == -1)
        MessageBox.Show("Escolha uma coluna numérica única para sorteio.", "Clique sob uma coluna numérica.");
      else
      {
        foreach (DataGridViewRow dr in dataGridView1.Rows)
          dr.DefaultCellStyle.BackColor = Color.White;
        Random r = new Random();
        int result = r.Next(MinNum, MaxNum + 1);

        foreach (DataGridViewRow dr in dataGridView1.Rows)
        {
          if (dr.Cells[ColIndex].Value != null)
          {
            int.TryParse(dr.Cells[ColIndex].Value.ToString(), out int i);
            if (result == i)
            {
              dr.DefaultCellStyle.BackColor = Color.LightGreen;


              try
              {
                dataGridView1.CurrentCell = dataGridView1.Rows[dr.Index].Cells[ColIndex];
              }
              catch (Exception ex)
              {
                Console.WriteLine("Erro ao selecionar a célula.", ex.ToString());
              }
              Thread.Sleep(100);
              MessageBox.Show(null, dr.Cells[ColIndex + 1].Value.ToString(), "Ganhou!", MessageBoxButtons.OK, MessageBoxIcon.Information);
              label1.Text = dr.Cells[ColIndex + 1].Value.ToString();
              break;
            }
          }
        }
      }
    }
  }
}
