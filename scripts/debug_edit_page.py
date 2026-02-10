"""
Debug script — opens the timesheet edit page and dumps the HTML structure
of the ribbon (BROWSE / TIMESHEET tabs) and the grid, so we can find
the correct selectors for Add Row and hour cells.

Run with: .venv/bin/python scripts/debug_edit_page.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bot.browser import BrowserManager
from bot.config import load_config
from bot.timesheet import TimesheetSummaryPage


def debug_edit_page():
    config = load_config()

    with BrowserManager(config) as bm:
        page = bm.page

        # Navigate to summary
        summary = TimesheetSummaryPage(page)
        summary.navigate()

        if bm.is_on_login_page(page):
            bm.wait_for_manual_login(page)
            summary.navigate()

        # Open current week's timesheet
        try:
            status = summary.open_timesheet()
            print(f"📋 Timesheet opened (was: {status})\n")
        except RuntimeError as e:
            print(f"❌ {e}")
            input("Press Enter to close...")
            return

        page.wait_for_timeout(3000)

        # =====================================================
        # 1. RIBBON — dump all tab links and buttons
        # =====================================================
        print("=" * 70)
        print("🎀 RIBBON STRUCTURE")
        print("=" * 70)

        # All anchor tags in the ribbon area
        ribbon = page.locator("#RibbonContainer, #s4-ribbonrow, [id*='Ribbon']")
        print(f"\n   Ribbon container count: {ribbon.count()}")

        # Find all <a> and <span> that look like tabs
        tabs = page.locator("a[id*='Tab'], li[id*='Tab']")
        print(f"\n   Tab-like elements: {tabs.count()}")
        for i in range(tabs.count()):
            el = tabs.nth(i)
            try:
                tag = el.evaluate("e => e.tagName")
                eid = el.get_attribute("id") or ""
                txt = el.inner_text(timeout=500).strip().replace("\n", " ")[:60]
                print(f"   [{i}] <{tag}> id=\"{eid}\"  text=\"{txt}\"")
            except Exception as ex:
                print(f"   [{i}] error: {ex}")

        # Also look for anything with "TIMESHEET" text
        print("\n   Elements containing 'TIMESHEET' text:")
        ts_els = page.locator("text=/TIMESHEET/i")
        for i in range(ts_els.count()):
            el = ts_els.nth(i)
            try:
                tag = el.evaluate("e => e.tagName")
                eid = el.get_attribute("id") or ""
                classes = el.get_attribute("class") or ""
                role = el.get_attribute("role") or ""
                txt = el.inner_text(timeout=500).strip().replace("\n", " ")[:80]
                print(f"   [{i}] <{tag}> id=\"{eid}\" class=\"{classes}\" role=\"{role}\" text=\"{txt}\"")
            except Exception as ex:
                print(f"   [{i}] error: {ex}")

        # =====================================================
        # 2. Click the TIMESHEET tab and dump ribbon buttons
        # =====================================================
        print("\n" + "=" * 70)
        print("🖱️  CLICKING TIMESHEET TAB...")
        print("=" * 70)

        # Try multiple selector strategies
        clicked = False
        strategies = [
            ("a with title containing Timesheet", "a[title*='Timesheet']"),
            ("a with id containing Timesheet", "a[id*='Timesheet'][id*='Tab']"),
            ("span exact text TIMESHEET", "span:text('TIMESHEET')"),
            ("any element with Timesheet text", "text=TIMESHEET"),
            ("li with Timesheet", "li[id*='Timesheet']"),
        ]
        for desc, sel in strategies:
            try:
                el = page.locator(sel).first
                if el.count() > 0:
                    el.click(timeout=2000)
                    page.wait_for_timeout(500)
                    print(f"   ✅ Clicked via: {desc}  (selector: {sel})")
                    clicked = True
                    break
            except Exception as ex:
                print(f"   ❌ {desc}: {ex}")

        if not clicked:
            print("   ⚠️  Could not click TIMESHEET tab via any strategy")

        # Now dump all visible ribbon buttons/commands
        print("\n   Ribbon buttons after clicking TIMESHEET tab:")
        buttons = page.locator(
            "#RibbonContainer a[id*='Button'], "
            "#RibbonContainer a[id*='Large'], "
            "#RibbonContainer span.ms-cui-ctl-largelabel, "
            "#RibbonContainer a.ms-cui-ctl-large, "
            "#RibbonContainer a.ms-cui-ctl-medium"
        )
        print(f"   Found {buttons.count()} ribbon buttons")
        for i in range(min(buttons.count(), 30)):
            el = buttons.nth(i)
            try:
                tag = el.evaluate("e => e.tagName")
                eid = el.get_attribute("id") or ""
                txt = el.inner_text(timeout=500).strip().replace("\n", " ")[:60]
                vis = el.is_visible()
                print(f"   [{i}] <{tag}> id=\"{eid}\"  text=\"{txt}\"  visible={vis}")
            except Exception as ex:
                print(f"   [{i}] error: {ex}")

        # Look specifically for "Add Row" or "Add Line"
        print("\n   Elements containing 'Add' text:")
        add_els = page.locator("text=/add row|add line|add task/i")
        for i in range(add_els.count()):
            el = add_els.nth(i)
            try:
                tag = el.evaluate("e => e.tagName")
                eid = el.get_attribute("id") or ""
                txt = el.inner_text(timeout=500).strip().replace("\n", " ")[:60]
                vis = el.is_visible()
                print(f"   [{i}] <{tag}> id=\"{eid}\"  text=\"{txt}\"  visible={vis}")
            except Exception as ex:
                print(f"   [{i}] error: {ex}")

        # =====================================================
        # 3. GRID — dump the task table structure
        # =====================================================
        print("\n" + "=" * 70)
        print("📊 GRID STRUCTURE")
        print("=" * 70)

        # Try to find the grid table
        grid_selectors = [
            "table[id*='GridData']",
            "table.MSPGridData",
            "#GridData",
            "div[id*='GridData'] table",
            "table[summary*='timesheet' i]",
            "table[summary*='grid' i]",
            "div[class*='grid' i] table",
            "table[id*='TSGrid']",
        ]
        for sel in grid_selectors:
            count = page.locator(sel).count()
            if count > 0:
                print(f"   ✅ Found: {sel} ({count} matches)")

        # Dump all tables and their IDs/classes
        tables = page.locator("table")
        print(f"\n   Total <table> elements: {tables.count()}")
        for i in range(tables.count()):
            t = tables.nth(i)
            try:
                tid = t.get_attribute("id") or ""
                tcls = t.get_attribute("class") or ""
                tsum = t.get_attribute("summary") or ""
                rows_count = t.locator("tr").count()
                if rows_count > 0 and (tid or tcls or tsum):
                    print(f"   table[{i}] id=\"{tid}\" class=\"{tcls[:60]}\" summary=\"{tsum[:40]}\" rows={rows_count}")
            except Exception:
                pass

        # Try finding rows with task names we know exist
        print("\n   Looking for known task names in the page...")
        known_tasks = ["ST-333", "Marketplace", "BAU", "Leave", "Enhancements"]
        for task in known_tasks:
            els = page.locator(f"text=/{task}/i")
            count = els.count()
            if count > 0:
                for j in range(min(count, 3)):
                    el = els.nth(j)
                    try:
                        tag = el.evaluate("e => e.tagName")
                        eid = el.get_attribute("id") or ""
                        parent_tag = el.evaluate("e => e.parentElement?.tagName || 'none'")
                        parent_id = el.evaluate("e => e.parentElement?.id || ''")
                        gp_tag = el.evaluate("e => e.parentElement?.parentElement?.tagName || 'none'")
                        gp_id = el.evaluate("e => e.parentElement?.parentElement?.id || ''")
                        print(f"   '{task}' [{j}] <{tag}> id=\"{eid}\" → parent <{parent_tag}> id=\"{parent_id}\" → gp <{gp_tag}> id=\"{gp_id}\"")
                    except Exception as ex:
                        print(f"   '{task}' [{j}] error: {ex}")

        # =====================================================
        # 4. INPUT CELLS — find all inputs/editable cells
        # =====================================================
        print("\n" + "=" * 70)
        print("✏️  EDITABLE CELLS")
        print("=" * 70)

        inputs = page.locator("input[type='text']")
        print(f"   <input type='text'>: {inputs.count()}")
        for i in range(min(inputs.count(), 15)):
            el = inputs.nth(i)
            try:
                eid = el.get_attribute("id") or ""
                val = el.get_attribute("value") or ""
                name = el.get_attribute("name") or ""
                vis = el.is_visible()
                print(f"   [{i}] id=\"{eid[:50]}\" name=\"{name[:30]}\" value=\"{val}\" visible={vis}")
            except Exception:
                pass

        ce = page.locator("td[contenteditable='true']")
        print(f"\n   <td contenteditable>: {ce.count()}")

        # Check for any other interactive grid elements
        divs_ce = page.locator("div[contenteditable='true']")
        print(f"   <div contenteditable>: {divs_ce.count()}")

        # =====================================================
        # 5. DUMP the grid area HTML (first grid-like structure)
        # =====================================================
        print("\n" + "=" * 70)
        print("🔬 RAW HTML AROUND TASK NAMES")
        print("=" * 70)
        try:
            # Find an element containing ST-333 and get its ancestor table's HTML
            st333 = page.locator("text=ST-333").first
            # Get the enclosing table row HTML
            row_html = st333.evaluate("""el => {
                let tr = el.closest('tr');
                if (tr) return tr.outerHTML.substring(0, 2000);
                // fallback: walk up 5 levels
                let node = el;
                for (let i = 0; i < 5; i++) {
                    if (node.parentElement) node = node.parentElement;
                }
                return node.outerHTML.substring(0, 2000);
            }""")
            print(row_html[:2000])
        except Exception as ex:
            print(f"   Error: {ex}")

        # Get the outerHTML of the parent table of ST-333
        print("\n" + "=" * 70)
        print("🔬 PARENT TABLE OF ST-333 (IDs and structure)")
        print("=" * 70)
        try:
            info = page.locator("text=ST-333").first.evaluate("""el => {
                let table = el.closest('table');
                if (!table) return 'no parent table found';
                let id = table.id || 'no-id';
                let cls = table.className || 'no-class';
                let rows = table.rows.length;
                let firstRowCells = table.rows[0] ? table.rows[0].cells.length : 0;
                
                // Get the IDs of all rows
                let rowInfo = [];
                for (let i = 0; i < Math.min(rows, 5); i++) {
                    let r = table.rows[i];
                    let cellTexts = [];
                    for (let j = 0; j < Math.min(r.cells.length, 4); j++) {
                        cellTexts.push(r.cells[j].innerText.substring(0, 30).trim());
                    }
                    rowInfo.push({
                        idx: i, 
                        id: r.id || '', 
                        cells: r.cells.length,
                        texts: cellTexts
                    });
                }
                return JSON.stringify({id, cls, rows, firstRowCells, rowInfo}, null, 2);
            }""")
            print(info)
        except Exception as ex:
            print(f"   Error: {ex}")

        # =====================================================
        # 6. DUMP hour cells structure around ST-333 row
        # =====================================================
        print("\n" + "=" * 70)
        print("🔬 HOUR CELLS FOR ST-333 ROW")
        print("=" * 70)
        try:
            info = page.locator("text=ST-333").first.evaluate("""el => {
                let tr = el.closest('tr');
                if (!tr) return 'no parent row';
                let cells = [];
                for (let i = 0; i < tr.cells.length; i++) {
                    let c = tr.cells[i];
                    let inputs = c.querySelectorAll('input');
                    let inputInfo = [];
                    inputs.forEach(inp => {
                        inputInfo.push({
                            type: inp.type,
                            id: inp.id.substring(0, 50),
                            name: inp.name.substring(0, 50),
                            value: inp.value
                        });
                    });
                    cells.push({
                        idx: i,
                        text: c.innerText.substring(0, 30).trim(),
                        tag: c.tagName,
                        id: (c.id || '').substring(0, 50),
                        cls: (c.className || '').substring(0, 50),
                        editable: c.contentEditable,
                        inputs: inputInfo
                    });
                }
                // Also check sibling rows (actual/planned)
                let next = tr.nextElementSibling;
                let nextInfo = null;
                if (next) {
                    let nc = [];
                    for (let i = 0; i < next.cells.length; i++) {
                        let c = next.cells[i];
                        let inputs = c.querySelectorAll('input');
                        let inputInfo = [];
                        inputs.forEach(inp => {
                            inputInfo.push({type: inp.type, id: inp.id.substring(0, 50), value: inp.value});
                        });
                        nc.push({
                            idx: i,
                            text: c.innerText.substring(0, 30).trim(),
                            id: (c.id || '').substring(0, 50),
                            cls: (c.className || '').substring(0, 50),
                            editable: c.contentEditable,
                            inputs: inputInfo
                        });
                    }
                    nextInfo = nc;
                }
                return JSON.stringify({taskRow: cells, nextRow: nextInfo}, null, 2);
            }""")
            print(info)
        except Exception as ex:
            print(f"   Error: {ex}")

        print("\n" + "=" * 70)
        input("Press Enter to close the browser...")

    print("👋 Done.")


if __name__ == "__main__":
    debug_edit_page()
