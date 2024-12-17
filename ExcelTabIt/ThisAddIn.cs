using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTabIt
{
    public partial class ThisAddIn
    {
        private Dictionary<Excel.Window, Microsoft.Office.Tools.CustomTaskPane> taskPanes;
        private List<Excel.Workbook> openWorkbooks = new List<Excel.Workbook>();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            taskPanes = new Dictionary<Excel.Window, Microsoft.Office.Tools.CustomTaskPane>();

            // Handle workbook open, close, and activate events
            this.Application.WorkbookOpen += WorkbookOpened;
            this.Application.WorkbookBeforeClose += WorkbookBeforeClosed;
            this.Application.WorkbookActivate += WorkbookActivated;

            // Initialize the task pane for the currently active workbook (if there is one)
            if (this.Application.Workbooks.Count > 0)
            {
                Excel.Workbook activeWorkbook = this.Application.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    ShowCustomTabsForWindow(activeWorkbook.Windows[1]);
                }
            }
        }

        private void WorkbookOpened(Excel.Workbook wb)
        {
            // Add the workbook to the list of open workbooks
            if (!openWorkbooks.Contains(wb))
            {
                openWorkbooks.Add(wb);
            }

            // Show the custom tabs for the new workbook's window
            ShowCustomTabsForWindow(wb.Windows[1]);
        }

        private void WorkbookBeforeClosed(Excel.Workbook wb, ref bool Cancel)
        {
            // Remove the workbook from the list of open workbooks
            openWorkbooks.Remove(wb);

            // Remove the task pane for the closing workbook's window
            RemoveCustomTabsForWindow(wb.Windows[1]);
        }

        private void WorkbookActivated(Excel.Workbook wb)
        {
            // Show or switch to the custom tabs for the active workbook's window
            ShowCustomTabsForWindow(wb.Windows[1]);
        }

        private void ShowCustomTabsForWindow(Excel.Window window)
        {
            // Ensure only one task pane per workbook window
            if (taskPanes.ContainsKey(window))
            {
                var existingTaskPane = taskPanes[window];
                if (existingTaskPane.Control is CustomTabControl existingTabControl)
                {
                    existingTabControl.ClearTabs(); // Clear existing tabs
                    foreach (var workbook in openWorkbooks)
                    {
                        existingTabControl.AddWorkbookTab(workbook); // Re-add the tabs
                    }

                    // Select the tab for the active workbook
                    existingTabControl.SelectTabForWorkbook(window.Application.ActiveWorkbook);
                }
                existingTaskPane.Visible = true; // Ensure it's visible
            }
            else
            {
                // No existing task pane, create a new one
                CustomTabControl tabControl = new CustomTabControl(this.Application);

                // Add workbook tabs for all open workbooks
                foreach (var workbook in openWorkbooks)
                {
                    tabControl.AddWorkbookTab(workbook);
                }

                // Create and show the new task pane
                var taskPane = this.CustomTaskPanes.Add(tabControl, "Open Tabs", window);
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                taskPane.Height = 60;
                taskPane.Visible = true;

                // Set the background color of the task pane to black
                taskPane.Control.BackColor = Color.Black;

                // Store the task pane in the dictionary
                taskPanes.Add(window, taskPane);

                // Select the tab for the active workbook
                tabControl.SelectTabForWorkbook(window.Application.ActiveWorkbook);
            }

            // Hide task panes for inactive workbooks
            foreach (var pane in taskPanes)
            {
                if (pane.Key != window)
                {
                    pane.Value.Visible = false;
                }
            }
        }

        private void RemoveCustomTabsForWindow(Excel.Window window)
        {
            if (taskPanes.ContainsKey(window))
            {
                // Remove the task pane and its reference from the dictionary
                this.CustomTaskPanes.Remove(taskPanes[window]);
                taskPanes.Remove(window);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Detach event handlers when shutting down
            this.Application.WorkbookOpen -= WorkbookOpened;
            this.Application.WorkbookBeforeClose -= WorkbookBeforeClosed;
            this.Application.WorkbookActivate -= WorkbookActivated;
        }
    }
}