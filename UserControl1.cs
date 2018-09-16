//Auhtor: Jhoel Perez
//Date: 6/25/2018

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using core = Microsoft.Office.Core;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
//add COM reference for microsoft office library and microsoft office excel


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
		// Contains a reference to the hosting application
		private excel.Application m_XlApplication = null;
		// Contains a reference to the active workbook
		private excel.Workbook m_Workbook = null;
		private WebBrowser webBrowser1;

		private readonly Missing MISS = Missing.Value;
		private bool m_ToolBarVisible = false;
		private Microsoft.Office.Core.CommandBar m_StandardCommandBar = null;
		/// <summary>Contains the path to the workbook file.</summary>
		private string m_ExcelFileName = string.Empty;
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

		public void file_open(string file)
		{
			if (!System.IO.File.Exists(file)) throw new Exception();
			file_name = file;
			//load work book in web browser
			//find make webbrowser 
			this.webBrowser1.Navigate(file, false);
		}

		public void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
		{
			//if workbook of the same file_name doesnt exist
			if ((m_Workbook = RetrieveWorkbook(file_name)) == null) return;
			//excel
			m_XlApplication = (Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
		}

		//pc executitnng programms list
		[DllImport("ole32.dll")]
		static extern int GetRunningObjectTable
						(uint reserved, out IRunningObjectTable pprot);
		[DllImport("ole32.dll")] static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);
		//fetch workbook
		public excel.Workbook RetrieveWorkbook(string xlfile)
		{
			IRunningObjectTable prot = null;
			IEnumMoniker pmonkenum = null;
			try
			{
				IntPtr pfetched = IntPtr.Zero;
				// Query the running object table (ROT)
				if (GetRunningObjectTable(0, out prot) != 0 || prot == null) return null;
				prot.EnumRunning(out pmonkenum); pmonkenum.Reset();
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

		private void AttachApplication()
		{
			try
			{
				if (m_ExcelFileName == null || m_ExcelFileName.Length == 0) return;
				// Creation of the workbook object
				if ((m_Workbook = RetrieveWorkbook(m_ExcelFileName)) == null) return;
				// Create the Excel.Application object
				m_XlApplication = (Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
				// Creation of the standard toolbar
				m_StandardCommandBar = m_XlApplication.CommandBars["Standard"];
				m_StandardCommandBar.Position = core.MsoBarPosition.msoBarTop;
				m_StandardCommandBar.Visible = m_ToolBarVisible;
				// Enable the OpenFile and New buttons
				foreach (core.CommandBarControl control in m_StandardCommandBar.Controls)
				{
					string name = control.get_accName(Missing.Value);
					if (name.Equals("Nouveau")) ((core.CommandBarButton)control).Enabled = false;
					if (name.Equals("Ouvrir")) ((core.CommandBarButton)control).Enabled = false;
				}
			}
			catch
			{
				MessageBox.Show("Impossible de charger le fichier Excel");
				return;
			}
		}
		#endregion
	}
}