//Auhtor: Jhoel Perez
//Date: 6/25/2018

using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;
using core = Microsoft.Office.Core;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
//add COM reference for microsoft office library and microsoft office excel
//xls to hmtl//Spire.xls.dll added to reference //Spire.Common.dll for save methods
using Spire.Xls;
using System.Data.OleDb;


namespace AddValue
{
	public partial class UserControl1 : UserControl
	{
		#region Contruction
		public UserControl1()
		{
			InitializeComponent();
		}
		#endregion Construction

		#region Fields
		private string file_name = "mm_sheet.xlsx";
		private string curr_sheet = "";
		// Contains a reference to the hosting application
		private excel.Application m_XlApplication = null;

		// Contains a reference to the active workbook
		private excel.Workbook m_Workbook = null;

		private readonly Missing MISS = Missing.Value;
		private bool m_ToolBarVisible = true;
		private Microsoft.Office.Core.CommandBar m_StandardCommandBar = null;
		///// <summary>Contains the path to the workbook file.</summary>
		//private string m_ExcelFileName = string.Empty;

		public bool ToolBarVisible { get; internal set; }
		#endregion Fields

		#region Events
		private void OnWebBrowserExcelNavigated(object sender, WebBrowserNavigatedEventArgs e)
		{
			AttachApplication();
		}

		private void OnToolBarVisibleChanged()
		{
			try
			{
				m_StandardCommandBar.Visible = m_ToolBarVisible;
			}
			catch { }
		}
		#endregion Events

		//#region Properties
		//[Browsable(false)]
		//public Workbook Workbook
		//{
		//	get { return m_Workbook; }
		//}

		//[Browsable(true), Category("Appearance")]
		//public bool ToolBarVisible
		//{
		//	get { return m_ToolBarVisible; }
		//	set
		//	{
		//		if (m_ToolBarVisible == value) return;
		//		m_ToolBarVisible = value;
		//		if (m_XlApplication != null) OnToolBarVisibleChanged();
		//	}
		//}
		//#endregion Properties

		#region Methods
		//public static DataTable ExcelToDataTable(string fileName)
		//{
		//	using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
		//	{
		//		using (var reader = ExcelReaderFactory.CreateReader(stream))
		//		{
		//			var result = reader.AsDataSet(new ExcelDataSetConfiguration()
		//			{
		//				UseColumnDataType = true,
		//				ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
		//				{
		//					UseHeaderRow = true
		//				}
		//			});
		//			DataTableCollection table = result.Tables;
		//			DataTable resultTable = table["Sheet1"];
		//			return resultTable;
		//		}
		//	}
		//}

		//Sheets names
		public ArrayList SheetList = new ArrayList();

		public void sheet_names(string file) {
			excel.Application objExcel = new excel.Application();
			Console.WriteLine(file);
			excel.Workbook Wb = objExcel.Workbooks.Open(file);
			excel.Worksheet Sheet = (excel.Worksheet)Wb.Sheets["Input Data"];

			foreach (excel.Worksheet objWorkSheets in Wb.Worksheets)
			{
				SheetList.Add(objWorkSheets.Name);
				Console.WriteLine(objWorkSheets.Name);
			}

			//xlWorkbook.SaveAs(temp_file_name, fileType, missing, missing, missing, missing,
			//		  Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
			//		  missing, missing, missing, missing, missing);
			//xlWorkbook.Close(false, missing, missing);
		}

		public void File_open(string file, DataGridView datagrid, string sheetName, ComboBox box1)
		{
			try
			{
				using (OleDbConnection conn = new OleDbConnection()) {

					DataTable sheet = new DataTable();
					//ask for sheet names from MS excel
					sheet_names(file);

					string fileExtension = Path.GetExtension(file);

					//What's Best file extension
					if (fileExtension == ".xlsm")
					{
						conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
					}

					using (OleDbCommand comm = new OleDbCommand())
					{
						//box1.Items.Equals(null)
						bool pass = false;
						bool done = false;

						//check input against sheet names 
						foreach (string tmp in SheetList) {
							//if (tmp.Equals(sheetName)){
								pass = true;
							//}
							box1.Items.Add(tmp);

							//if full
							if (!sheetName.Equals("") || sheetName.Equals(""))
							{
								sheetName = tmp;
								done = true;
							}

							//if (sheetName.Equals(tmp)) {
							//	Console.WriteLine("Found the Sheet");
							//	sheetName = tmp;
							//}
						}

						
						if (true){
								comm.CommandText = "Select * from [" + sheetName + "$]";

								comm.Connection = conn;

								using (OleDbDataAdapter da = new OleDbDataAdapter())
								{
									da.SelectCommand = comm;
									da.Fill(sheet);
								//assign sheet
								//datagrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
								datagrid.DataSource = sheet;


								}
							}
							else {
								throw new System.ArgumentException("Invalid sheet name", "original");
							}
						
					}
				}
			}

			catch (IOException e)
			{
				MessageBox.Show("Invalid excel format, try again!");
				if (e.Source != null)
					Console.WriteLine("IOException source: {0}", e.Source);
			}
			catch (System.Exception ex ) {
				MessageBox.Show("Invalid Sheet name, try again!");
				if (ex.Source != null)
					Console.WriteLine("IOException source: {0}", ex.Source);
			}

			return;
			//load work book in web browser
			//find make webbrowser 
			//Console.WriteLine("Navigate Web Browser");

			//excel pop up on navigate method
			//(this.webBrowser1).Navigate(new Uri(@file), false);
		}

		//WEB BROWSER
		string GetTempFile(string extension)
		{

			Uri baseUri = new Uri(Path.GetTempPath());
			Uri myUri = new Uri(baseUri, Path.ChangeExtension(Path.GetRandomFileName(), extension));
			// Uses the Combine, GetTempPath, ChangeExtension, 
			// and GetRandomFile methods of Path to 
			// create a temp file of the extension we're looking for. 
			//return Path.Combine(Path.GetTempPath(),	Path.ChangeExtension(Path.GetRandomFileName(), extension));
			return (myUri.AbsolutePath);
		}

		public void convert_excel(string file_name)
		{
			Console.WriteLine("conver excel mathod ");
			string dir = Directory.GetCurrentDirectory();
			try
			{
				//spire workbook
				//Workbook spire_workbook = new Workbook();
				//Console.WriteLine(file_name);
				//spire_workbook.LoadFromFile(file_name);
				//Worksheet spire_sheet = spire_workbook.Worksheets[0];
				//spire_sheet.SaveToHtml("excel.html");

				//Path.Combine(dir,"excel.html");
				//Console.WriteLine("path to html is {0}", dir);
				//webBrowser1.Url = new Uri(dir);
				
				//System.Diagnostics.Process.Start(temp_html);

				excel.Application excel = new excel.Application();
				excel.Visible = false;
				excel.Workbook xlWorkbook = excel.Workbooks.Open(file_name);
				//excel.Worksheet xlWorksheet = xlWorkbook.Sheets[nameOfSheet];
				excel.Worksheet xlWorksheet = xlWorkbook.Sheets[0];

				string temp_file_name = GetTempFile("html");

				object missing = System.Reflection.Missing.Value;
				object newFileName = (object)temp_file_name;
				object fileType = (object)excel.GetType();

				xlWorkbook.SaveAs(temp_file_name, fileType, missing, missing, missing, missing,
					  Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
					  missing, missing, missing, missing, missing);
				xlWorkbook.Close(false, missing, missing);
			}

			catch
			{
				Console.WriteLine("Couldnt convert xmls to html");
			}
		}

		public void File_open_web(string file)
		{
			Console.WriteLine(file);
			try
			{
				if (!System.IO.File.Exists(file))
				{
					throw new Exception();
				}
				else
				{
					//file exists
					//load work book in web browser
					//find make webbrowser 
					//Console.WriteLine("Navigate Web Browser");
					load_sheet(file);
				}
			}

			catch
			{
				MessageBox.Show("Invalid excel format, try again!");
				return;
			}

		}

		public void load_sheet(string fileName)
		{
			// Call ConvertDocument asynchronously. 
			Console.WriteLine("Load sheet on web");
			//Console.WriteLine("Asking to use Input Data sheet from model");
			//string sheet_name = "Input Data";
			//convert_excel(fileName);
			Console.WriteLine("file path ");
			(this.webBrowser1).Navigate(@file_name,false);
		}
		//END WEB BROWSER

		public void WebBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
		{
			Console.WriteLine("web browser 1 navigated");
			//if workbook of the same file_name doesnt exist
			Console.WriteLine("file_name %s",file_name);
			if ((m_Workbook = RetrieveWorkbook(file_name)) == null) return;
			//excel
			m_XlApplication = (Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
		}

		public excel.Workbook GetActiveWorkbook(string xlfile)
		{
			IRunningObjectTable prot = null;
			IEnumMoniker pmonkenum = null;
			try
			{
				IntPtr pfetched = IntPtr.Zero;
				// Query the running object table (ROT)
				if (GetRunningObjectTable(0, out prot) != 0 || prot == null) return null;
				prot.EnumRunning(out pmonkenum);
				pmonkenum.Reset();
				IMoniker[] monikers = new IMoniker[1];
				while (pmonkenum.Next(1, monikers, pfetched) == 0)
				{
					IBindCtx pctx; string filepathname;
					CreateBindCtx(0, out pctx);
					// Get the name of the file
					monikers[0].GetDisplayName(pctx, null, out filepathname);
					// Clean up
					Marshal.ReleaseComObject(pctx);
					// Search for the workbook
					if (filepathname.IndexOf(xlfile) != -1)
					{
						object roval;
						// Get a handle on the workbook
						prot.GetObject(monikers[0], out roval);
						return roval as excel.Workbook;
					}
				}
			}
			catch
			{
				return null;
			}
			finally
			{
				// Clean up
				if (prot != null) Marshal.ReleaseComObject(prot);
				if (pmonkenum != null) Marshal.ReleaseComObject(pmonkenum);
			}
			return null;
		}


		//pc executitnng programms list
		[DllImport("ole32.dll")]
		static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);
		[DllImport("ole32.dll")] static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

		//FAILS////////////////////////////////////////////////////////////////////////////

		//
		//fetch workbook
		public excel.Workbook RetrieveWorkbook(string xlfile)
		{
			Console.WriteLine("START Retrieve Workbook");
			IRunningObjectTable prot = null;
			IEnumMoniker pmonkenum = null;
			try
			{
				IntPtr pfetched = IntPtr.Zero;
				// Query the running object table (ROT)
				if (GetRunningObjectTable(0, out prot) != 0 || prot == null) {
					Console.WriteLine("Retrieved object NULL");
					return null;
				}
				prot.EnumRunning(out pmonkenum); pmonkenum.Reset();
				IMoniker[] monikers = new IMoniker[1];
				while (pmonkenum.Next(1, monikers, pfetched) == 0)
				{
					IBindCtx pctx;
					string filepathname;
					CreateBindCtx(0, out pctx);
					// Get the name of the file
					monikers[0].GetDisplayName(pctx, null, out filepathname);
					// Clean up
					Marshal.ReleaseComObject(pctx);
					// Search for the workbook
					if (filepathname.IndexOf(xlfile) != -1)
					{
						object roval;
						// Get a handle on the workbook
						prot.GetObject(monikers[0], out roval);
						//good return type
						return roval as excel.Workbook;
					}
				}
			}
			catch
			{
				MessageBox.Show("Excel retrieve workbook failed");
				return null;
			}
			finally
			{
				//if (m_Workbook == null) {
				//	Console.WriteLine("failed to retrieve workbook");
				//}
				// Clean up
				if (prot != null) Marshal.ReleaseComObject(prot);
				if (pmonkenum != null) Marshal.ReleaseComObject(pmonkenum);
			}

			Console.WriteLine("END Retrieve Workbook");
			return null;
		}

		protected void OnClosed(object sender, EventArgs e)
		{
			try
			{
				// Quit Excel and clean up.
				if (m_Workbook != null)
				{
					m_Workbook.Close(true, Missing.Value, Missing.Value);
					System.Runtime.InteropServices.Marshal.ReleaseComObject
											(m_Workbook);
					m_Workbook = null;
				}
				if (m_XlApplication != null)
				{
					m_XlApplication.Quit();
					System.Runtime.InteropServices.Marshal.ReleaseComObject
										(m_XlApplication);
					m_XlApplication = null;
					System.GC.Collect();
				}
			}
			catch
			{
				MessageBox.Show("Failed to close the application");
			}
		}

		private excel.Workbook Open_Workbook(string file_path)
		{
			excel.Workbook save_copy = null;
			try
			{
				if (m_XlApplication != null)
				{
					save_copy = m_XlApplication.Workbooks.Open(file_path, 0, true);
				}
				else {
					Console.WriteLine("m_XLApplication not defined! ");
				}
				return save_copy;
			}
			catch (Exception e)
			{
				throw new ArgumentException("Error opening %",file_path, e);
			}
		}

		private void AttachApplication()
		{
			try
			{
				Console.WriteLine("USerControl1.cs AttachApplication");
				
				if (file_name == null || file_name.Length == 0) {
					Console.WriteLine("Error: not valid file name provided", file_name);
					return;
				}
				Console.WriteLine("file name is: ");
				Console.WriteLine(file_name);

				//m_Workbook = GetActiveWorkbook(file_name);
				m_Workbook = RetrieveWorkbook(file_name);
				// Creation of the workbook object
				int count = 0;
				while (m_Workbook  == null && count <5) {
					Console.WriteLine("Couldn't retrieve workbook %",file_name);
					count++;
					m_Workbook = Open_Workbook(file_name);
				}

				if (m_Workbook != null) {
					Console.WriteLine("Workbook not null");
					// Create the Excel.Application object
					m_XlApplication = (excel.Application)m_Workbook.Application;
					// Creation of the standard toolbar
					m_StandardCommandBar = m_XlApplication.CommandBars["Standard"];
					m_StandardCommandBar.Position = core.MsoBarPosition.msoBarTop;
					m_StandardCommandBar.Visible = m_ToolBarVisible;
					//value may be false
					Console.WriteLine("tool bar bool: ");
					Console.WriteLine(m_ToolBarVisible);
					// Enable the OpenFile and New buttons
					foreach (core.CommandBarControl control in m_StandardCommandBar.Controls)
					{
						string name = control.get_accName(Missing.Value);
						if (name.Equals("Nouveau")) ((core.CommandBarButton)control).Enabled = false;
						if (name.Equals("Ouvrir")) ((core.CommandBarButton)control).Enabled = false;
					}
				}
				
			}
			catch
			{
				MessageBox.Show("Impossible to load Excel file");
				return;
			}
			Console.WriteLine("END USerControl1.cs AttachApplication");
		}
		#endregion
	}
}