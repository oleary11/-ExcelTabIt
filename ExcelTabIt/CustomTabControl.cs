using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTabIt
{
    public partial class CustomTabControl : UserControl
    {
        private List<Excel.Workbook> openWorkbooks;
        private Excel.Application excelApp;

        // Define the TabControl for the tabs
        private TabControl tabControl;

        // Variable to track hovered tab index
        private int hoveredTabIndex = -1;

        public CustomTabControl(Excel.Application excel)
        {
            InitializeComponent();
            openWorkbooks = new List<Excel.Workbook>();
            excelApp = excel;

            // Initialize the TabControl with custom drawing
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                DrawMode = TabDrawMode.OwnerDrawFixed,
                ItemSize = new Size(100, 30),  // Set the size of the tabs
                Alignment = TabAlignment.Top, // Tabs positioned at the top
                SizeMode = TabSizeMode.Fixed, // Ensure fixed sizing
                BackColor = Color.Black // Set TabControl background to black
            };

            // Override Paint event to ensure background is black
            this.Paint += CustomTabControl_Paint;

            // Subscribe to the draw item event for custom rendering
            tabControl.DrawItem += TabControl_DrawItem;
            tabControl.MouseMove += TabControl_MouseMove;

            tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

            // Add the TabControl to the user control
            this.Controls.Add(tabControl);
        }

        // Override the Paint event to enforce black background
        private void CustomTabControl_Paint(object sender, PaintEventArgs e)
        {
            this.BackColor = Color.Black;
            e.Graphics.FillRectangle(Brushes.Black, this.ClientRectangle);
        }

        private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            TabPage tabPage = tabControl.TabPages[e.Index];
            Rectangle tabBounds = tabControl.GetTabRect(e.Index);

            // Fill entire tab area with black background
            g.FillRectangle(Brushes.Black, tabBounds);

            // Draw green bar for selected or hovered tabs
            if (e.State == DrawItemState.Selected)
            {
                // Bright green bar for selected tab
                g.FillRectangle(new SolidBrush(Color.FromArgb(0, 255, 127)), tabBounds.X, tabBounds.Bottom - 5, tabBounds.Width, 5);
            }
            else if (e.Index == hoveredTabIndex)
            {
                // Faded green bar for hovered tab
                g.FillRectangle(new SolidBrush(Color.FromArgb(0, 176, 80)), tabBounds.X, tabBounds.Bottom - 5, tabBounds.Width, 5);
            }

            // Draw tab text centered in white, without borders
            TextRenderer.DrawText(g, tabPage.Text, tabControl.Font, tabBounds, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void TabControl_MouseMove(object sender, MouseEventArgs e)
        {
            // Track hovered tab to trigger redraw for hover effect
            for (int i = 0; i < tabControl.TabCount; i++)
            {
                Rectangle tabRect = tabControl.GetTabRect(i);
                if (tabRect.Contains(e.Location))
                {
                    hoveredTabIndex = i;
                    tabControl.Invalidate(); // Trigger redraw for hover effect
                    return;
                }
            }

            // Reset hover effect if mouse leaves the tab area
            hoveredTabIndex = -1;
            tabControl.Invalidate();
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab != null)
            {
                // Find the workbook that corresponds to the selected tab and activate it
                var selectedTabName = tabControl.SelectedTab.Text;
                foreach (var wb in openWorkbooks)
                {
                    if (wb.Name == selectedTabName)
                    {
                        // Activate the workbook
                        wb.Activate();

                        // Bring the workbook's window to the foreground
                        if (wb.Windows.Count > 0)
                        {
                            wb.Windows[1].Activate(); // Activate the first window of the workbook
                        }
                        break;
                    }
                }
            }
        }

        public void AddWorkbookTab(Excel.Workbook wb)
        {
            // Create a new tab for the workbook (no sheet names)
            TabPage tabPage = new TabPage
            {
                Text = wb.Name // Only the workbook name, no sheets
            };

            // Add the tab to the TabControl
            tabControl.TabPages.Add(tabPage);

            // Keep track of the workbook
            openWorkbooks.Add(wb);
        }

        // Method to select the tab for a specific workbook
        public void SelectTabForWorkbook(Excel.Workbook wb)
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Text == wb.Name)
                {
                    // Select the tab corresponding to the active workbook
                    tabControl.SelectedTab = tabPage;
                    break;
                }
            }
        }

        // Method to remove the workbook tab
        public void RemoveWorkbookTab(Excel.Workbook wb)
        {
            // Find and remove the corresponding tab
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Text == wb.Name)
                {
                    tabControl.TabPages.Remove(tabPage);
                    break;
                }
            }

            openWorkbooks.Remove(wb);
        }

        // Method to clear all tabs
        public void ClearTabs()
        {
            tabControl.TabPages.Clear();
            openWorkbooks.Clear();
        }
    }
}