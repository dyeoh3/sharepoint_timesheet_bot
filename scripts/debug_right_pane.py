"""
Debug script — inspects the RIGHT PANE of the timesheet grid (hour cells)
and tries clicking cells to see how they become editable.

Run with: .venv/bin/python scripts/debug_right_pane.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import load_config
from bot.timesheet import TimesheetSummaryPage


def debug_right_pane():
    config = load_config()

    with BrowserManager(config) as bm:
        page = bm.page

        summary = TimesheetSummaryPage(page)
        summary.navigate()

        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            summary.navigate()

        try:
            status = summary.open_timesheet()
            print(f"📋 Timesheet opened (was: {status})\n")
        except RuntimeError as e:
            print(f"❌ {e}")
            input("Press Enter to close...")
            return

        page.wait_for_timeout(3000)

        # =====================================================
        # 1. RIGHT PANE TABLE structure
        # =====================================================
        right_pane_id = "ctl00_ctl00_ctl32_g_baa84582_ea68_426d_8fea_fc4ba2b73718_TimesheetPartJSGridControl_rightpane_mainTable"
        
        print("=" * 70)
        print("📊 RIGHT PANE TABLE (hour cells)")
        print("=" * 70)
        
        info = page.evaluate(f"""() => {{
            let table = document.getElementById('{right_pane_id}');
            if (!table) return 'Right pane table not found';
            
            let rows = [];
            for (let i = 0; i < table.rows.length; i++) {{
                let r = table.rows[i];
                let cells = [];
                for (let j = 0; j < r.cells.length; j++) {{
                    let c = r.cells[j];
                    cells.push({{
                        idx: j,
                        text: c.innerText.substring(0, 20).trim(),
                        cls: (c.className || '').substring(0, 40),
                        colspan: c.colSpan,
                        editable: c.contentEditable,
                        role: c.getAttribute('role') || '',
                        childTags: Array.from(c.children).map(ch => ch.tagName).join(',')
                    }});
                }}
                rows.push({{
                    idx: i, 
                    id: r.id || '',
                    role: r.getAttribute('role') || '',
                    cellCount: r.cells.length,
                    cells: cells
                }});
            }}
            return JSON.stringify(rows, null, 2);
        }}""")
        print(info)

        # =====================================================
        # 2. Map left pane rows to right pane rows
        # =====================================================
        left_pane_id = "ctl00_ctl00_ctl32_g_baa84582_ea68_426d_8fea_fc4ba2b73718_TimesheetPartJSGridControl_leftpane_mainTable"
        
        print("\n" + "=" * 70)
        print("🔗 LEFT/RIGHT PANE ROW MAPPING")
        print("=" * 70)
        
        mapping = page.evaluate(f"""() => {{
            let left = document.getElementById('{left_pane_id}');
            let right = document.getElementById('{right_pane_id}');
            if (!left || !right) return 'Tables not found';
            
            let result = [];
            for (let i = 0; i < left.rows.length; i++) {{
                let lr = left.rows[i];
                let taskCell = lr.cells.length >= 3 ? lr.cells[2] : null;
                let taskName = taskCell ? taskCell.innerText.substring(0, 35).trim() : '';
                
                // Find corresponding right pane rows
                // Usually: right has 2 rows per left row (Actual + Planned)
                // Header maps to header
                let rightIdx1 = (i === 0) ? 0 : (i - 1) * 2 + 1;
                let rightIdx2 = rightIdx1 + 1;
                
                let rr1 = right.rows[rightIdx1];
                let rr2 = right.rows[rightIdx2];
                
                result.push({{
                    leftIdx: i,
                    taskName: taskName,
                    rightIdx1: rightIdx1,
                    rightRow1FirstCell: rr1 ? rr1.cells[0]?.innerText?.substring(0, 20).trim() : 'N/A',
                    rightRow1Cells: rr1 ? rr1.cells.length : 0,
                    rightIdx2: rightIdx2,
                    rightRow2FirstCell: rr2 ? rr2.cells[0]?.innerText?.substring(0, 20).trim() : 'N/A',
                    rightRow2Cells: rr2 ? rr2.cells.length : 0,
                }});
            }}
            return JSON.stringify(result, null, 2);
        }}""")
        print(mapping)

        # =====================================================
        # 3. Click a cell in the ST-333 "Actual" row to activate it
        # =====================================================
        print("\n" + "=" * 70)
        print("🖱️  CLICKING A CELL TO ACTIVATE IT")
        print("=" * 70)
        
        # First, find which right-pane row index corresponds to ST-333
        st333_info = page.evaluate(f"""() => {{
            let left = document.getElementById('{left_pane_id}');
            let right = document.getElementById('{right_pane_id}');
            if (!left || !right) return 'Tables not found';
            
            // Find ST-333 in left pane
            let st333LeftIdx = -1;
            for (let i = 0; i < left.rows.length; i++) {{
                let taskCell = left.rows[i].cells.length >= 3 ? left.rows[i].cells[2] : null;
                if (taskCell && taskCell.innerText.includes('ST-333')) {{
                    st333LeftIdx = i;
                    break;
                }}
            }}
            
            if (st333LeftIdx < 0) return 'ST-333 not found in left pane';
            
            // Actual row in right pane
            let actualIdx = (st333LeftIdx - 1) * 2 + 1;
            let plannedIdx = actualIdx + 1;
            
            let actualRow = right.rows[actualIdx];
            let plannedRow = right.rows[plannedIdx];
            
            if (!actualRow) return 'Actual row not found at index ' + actualIdx;
            
            let cellDetails = [];
            for (let j = 0; j < actualRow.cells.length; j++) {{
                let c = actualRow.cells[j];
                cellDetails.push({{
                    idx: j,
                    text: c.innerText.substring(0, 20).trim(),
                    cls: (c.className || '').substring(0, 60),
                    width: c.offsetWidth,
                    role: c.getAttribute('role') || '',
                    style: c.style.cssText.substring(0, 60),
                    editable: c.contentEditable,
                    html: c.innerHTML.substring(0, 100)
                }});
            }}
            
            return JSON.stringify({{
                st333LeftIdx,
                actualIdx,
                plannedIdx,
                actualRowCells: actualRow.cells.length,
                cellDetails
            }}, null, 2);
        }}""")
        print(st333_info)

        # Now actually click the first day cell and see what happens
        print("\n   Clicking first day cell in ST-333 Actual row...")
        
        click_result = page.evaluate(f"""() => {{
            let left = document.getElementById('{left_pane_id}');
            let right = document.getElementById('{right_pane_id}');
            
            let st333LeftIdx = -1;
            for (let i = 0; i < left.rows.length; i++) {{
                let taskCell = left.rows[i].cells.length >= 3 ? left.rows[i].cells[2] : null;
                if (taskCell && taskCell.innerText.includes('ST-333')) {{
                    st333LeftIdx = i;
                    break;
                }}
            }}
            
            let actualIdx = (st333LeftIdx - 1) * 2 + 1;
            let actualRow = right.rows[actualIdx];
            
            // The first cell is usually "Time Type" label, actual day cells start after
            // Find the first cell that looks like a day cell (usually index 1+)
            let targetCell = null;
            for (let j = 1; j < actualRow.cells.length; j++) {{
                let c = actualRow.cells[j];
                let role = c.getAttribute('role') || '';
                if (role === 'gridcell') {{
                    targetCell = c;
                    break;
                }}
            }}
            
            if (!targetCell) {{
                // Fallback: just pick cell at index 1
                targetCell = actualRow.cells[1];
            }}
            
            if (!targetCell) return 'No target cell found';
            
            // Click it
            targetCell.click();
            
            return JSON.stringify({{
                clickedCellIdx: targetCell.cellIndex,
                text: targetCell.innerText.substring(0, 20),
                html: targetCell.innerHTML.substring(0, 200)
            }});
        }}""")
        print(f"   Click result: {click_result}")
        
        page.wait_for_timeout(1000)
        
        # Now check if any inputs appeared
        print("\n   After click — checking for inputs...")
        post_click = page.evaluate(f"""() => {{
            let right = document.getElementById('{right_pane_id}');
            let inputs = right.querySelectorAll('input[type="text"]');
            let editables = right.querySelectorAll('[contenteditable="true"]');
            let allInputs = right.querySelectorAll('input');
            
            let inputDetails = [];
            allInputs.forEach(inp => {{
                inputDetails.push({{
                    type: inp.type,
                    id: inp.id.substring(0, 50),
                    value: inp.value,
                    visible: inp.offsetParent !== null,
                    cls: (inp.className || '').substring(0, 50)
                }});
            }});
            
            // Also check for active element
            let active = document.activeElement;
            let activeInfo = {{
                tag: active?.tagName,
                id: (active?.id || '').substring(0, 50),
                cls: (active?.className || '').substring(0, 50),
                type: active?.type || '',
                editable: active?.contentEditable
            }};
            
            return JSON.stringify({{
                textInputs: inputs.length,
                editables: editables.length,
                allInputs: inputDetails,
                activeElement: activeInfo
            }}, null, 2);
        }}""")
        print(post_click)

        # Try double-clicking instead
        print("\n   Trying double-click on the same cell...")
        page.evaluate(f"""() => {{
            let left = document.getElementById('{left_pane_id}');
            let right = document.getElementById('{right_pane_id}');
            
            let st333LeftIdx = -1;
            for (let i = 0; i < left.rows.length; i++) {{
                let taskCell = left.rows[i].cells.length >= 3 ? left.rows[i].cells[2] : null;
                if (taskCell && taskCell.innerText.includes('ST-333')) {{
                    st333LeftIdx = i;
                    break;
                }}
            }}
            
            let actualIdx = (st333LeftIdx - 1) * 2 + 1;
            let actualRow = right.rows[actualIdx];
            let cell = actualRow.cells[1];
            
            let dblclick = new MouseEvent('dblclick', {{bubbles: true, cancelable: true}});
            cell.dispatchEvent(dblclick);
        }}""")
        
        page.wait_for_timeout(1000)
        
        post_dbl = page.evaluate(f"""() => {{
            let right = document.getElementById('{right_pane_id}');
            let inputs = right.querySelectorAll('input[type="text"]');
            let editables = right.querySelectorAll('[contenteditable="true"]');
            
            let active = document.activeElement;
            
            // Check if the cell now has an input
            let left = document.getElementById('{left_pane_id}');
            let st333LeftIdx = -1;
            for (let i = 0; i < left.rows.length; i++) {{
                let taskCell = left.rows[i].cells.length >= 3 ? left.rows[i].cells[2] : null;
                if (taskCell && taskCell.innerText.includes('ST-333')) {{
                    st333LeftIdx = i;
                    break;
                }}
            }}
            let actualIdx = (st333LeftIdx - 1) * 2 + 1;
            let actualRow = right.rows[actualIdx];
            let cell = actualRow.cells[1];
            
            return JSON.stringify({{
                textInputs: inputs.length,
                editables: editables.length,
                cellHTML: cell.innerHTML.substring(0, 300),
                activeTag: active?.tagName,
                activeId: (active?.id || '').substring(0, 50),
                activeType: active?.type || ''
            }}, null, 2);
        }}""")
        print(f"   After double-click: {post_dbl}")

        # Try using Playwright's click method on the actual cell locator
        print("\n   Trying Playwright dblclick on cell...")
        right_table = page.locator(f"#{right_pane_id}")
        right_rows = right_table.locator("tr")
        print(f"   Right pane rows: {right_rows.count()}")
        
        # Find ST-333 actual row (we need to figure out the mapping)
        # From the left pane, ST-333 was at index 7 (0=header, tasks start at 1)
        # Right pane actual row = (7-1)*2 + 1 = 13
        # Let's verify by checking all right pane row texts
        for i in range(right_rows.count()):
            r = right_rows.nth(i)
            first_cell_text = ""
            try:
                first_cell = r.locator("td").first
                first_cell_text = first_cell.inner_text(timeout=500).strip()
            except:
                pass
            if first_cell_text:
                print(f"   Right row {i}: \"{first_cell_text}\"")
        
        print("\n" + "=" * 70)
        input("Press Enter to close the browser...")

    print("👋 Done.")


if __name__ == "__main__":
    debug_right_pane()
