/*
 * Created by SharpDevelop.
 * User: galiev.re
 * Date: 07.06.2017
 * Time: 18:50
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace newexcel
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
		

//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void Button1Click(object sender, EventArgs e)
		{
		Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
		Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
		Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
		//Книга.
		ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
		//Таблица.
		ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
		
		//Значения [y - строка,x - столбец]
		ObjWorkSheet.Cells[1, 1] = textBox1.Text;
		ObjWorkSheet.Cells[1, 2] = textBox2.Text;
		ObjWorkSheet.Cells[1, 3] = textBox3.Text;
		ObjExcel.Visible = true;
		ObjExcel.UserControl = true;	
		}
	}
}
