import { ReactElement, createElement, useState, useCallback, useEffect, useRef } from "react";
import classNames from "classnames";
import { TableviewContainerProps } from "../typings/TableviewProps";
import Big from "big.js";
import "./ui/Tableview.css";

// ── Interfaces ────────────────────────────────────────────────────────────────
interface CellObject {
    id: string;
    sequenceNumber: string;
    isBlocked: boolean;
    isMerged: boolean;
    mergeId: string;
    isBlank: boolean;
    rowIndex: number;
    columnIndex: number;
    checked: boolean;
    isSelected: boolean;
    rowSpan: number;
    colSpan: number;
    isHidden: boolean;
}

interface TableRow {
    id: string;
    rowIndex: number;
    cells: CellObject[];
}

interface TableData {
    rows: number;
    columns: number;
    tableRows: TableRow[];
    metadata?: { createdAt?: string; updatedAt?: string };
}

// ── Pure helpers (module-level, no hooks) ─────────────────────────────────────

// Walk up DOM from any element to find nearest <td data-cellid="...">
function getTdCell(el: Element | null): HTMLElement | null {
    let cur = el;
    while (cur) {
        if (cur.tagName === "TD" && (cur as HTMLElement).dataset.cellid)
            return cur as HTMLElement;
        cur = cur.parentElement;
    }
    return null;
}

// Parse "cell_R_C" → {row, col}
function parseCellId(id: string): { row: number; col: number } | null {
    const p = id.replace("cell_", "").split("_");
    if (p.length < 2) return null;
    const row = parseInt(p[0], 10);
    const col = parseInt(p[1], 10);
    if (isNaN(row) || isNaN(col)) return null;
    return { row, col };
}

// Build a Set of all cell IDs in the rectangle defined by two corners
function rectSet(r1: number, c1: number, r2: number, c2: number): Set<string> {
    const s = new Set<string>();
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++)
            s.add(`cell_${r}_${c}`);
    return s;
}

// ── Component ─────────────────────────────────────────────────────────────────
const Tableview = (props: TableviewContainerProps): ReactElement => {

    // ── Initial counts from Mendix attributes ─────────────────────────────────
    const initRows = () => {
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value)
            return Math.max(1, Number(props.rowCountAttribute.value));
        return 3;
    };
    const initCols = () => {
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value)
            return Math.max(1, Number(props.columnCountAttribute.value));
        return 3;
    };

    // ── State ─────────────────────────────────────────────────────────────────
    const [rowCount,      setRowCount]      = useState<number>(initRows());
    const [columnCount,   setColumnCount]   = useState<number>(initCols());
    const [tableRows,     setTableRows]     = useState<TableRow[]>([]);
    const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
    const [isDragging,    setIsDragging]    = useState<boolean>(false);
    const [isInitialLoad, setIsInitialLoad] = useState<boolean>(true);
    const [isSaving,      setIsSaving]      = useState<boolean>(false);
    const [dataLoaded,    setDataLoaded]    = useState<boolean>(false);

    // ── Stable refs (mutations don't trigger renders) ─────────────────────────
    const lastSavedRef          = useRef<string>("");
    const isUserInputRef        = useRef<boolean>(false);
    const ignoreAttrUpdateRef   = useRef<boolean>(false);
    const latestStateRef        = useRef<{ rows: TableRow[]; rc: number; cc: number } | null>(null);
    const saveTimerRef          = useRef<NodeJS.Timeout | null>(null);

    // Drag refs — written/read only in event handlers, never trigger renders
    const dragging              = useRef<boolean>(false);
    const dragStart             = useRef<{ row: number; col: number } | null>(null);
    const dragMoved             = useRef<boolean>(false);

    // Double-click detection
    const lastClickTime         = useRef<number>(0);
    const lastClickId           = useRef<string>("");

    // tableRows ref — always holds current tableRows without stale closure issues
    const tableRowsRef          = useRef<TableRow[]>([]);
    tableRowsRef.current = tableRows;

    // ── Statistics ────────────────────────────────────────────────────────────
    const updateStats = useCallback((rows: TableRow[]) => {
        const total   = rows.reduce((s, r) => s + r.cells.length, 0);
        const blocked = rows.reduce((s, r) => s + r.cells.filter(c => c.isBlocked).length, 0);
        const merged  = rows.reduce((s, r) => s + r.cells.filter(c => c.isMerged && !c.isHidden).length, 0);
        if (props.totalCellsAttribute?.status   === "available") props.totalCellsAttribute.setValue(new Big(total));
        if (props.blockedCellsAttribute?.status === "available") props.blockedCellsAttribute.setValue(new Big(blocked));
        if (props.mergedCellsAttribute?.status  === "available") props.mergedCellsAttribute.setValue(new Big(merged));
    }, [props.totalCellsAttribute, props.blockedCellsAttribute, props.mergedCellsAttribute]);

    // ── Save ──────────────────────────────────────────────────────────────────
    const saveToBackend = useCallback((rows: TableRow[], rc: number, cc: number) => {
        const json = JSON.stringify({ rows: rc, columns: cc, tableRows: rows, metadata: { updatedAt: new Date().toISOString() } });
        lastSavedRef.current = json;
        setIsSaving(true);
        if (props.useAttributeData?.status   === "available") props.useAttributeData.setValue(json);
        if (props.tableDataAttribute?.status === "available") props.tableDataAttribute.setValue(json);
        ignoreAttrUpdateRef.current = true;
        if (props.rowCountAttribute?.status    === "available") props.rowCountAttribute.setValue(new Big(rc));
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(cc));
        updateStats(rows);
        if (props.onTableChange?.canExecute) props.onTableChange.execute();
        setTimeout(() => setIsSaving(false), 200);
    }, [props.useAttributeData, props.tableDataAttribute, props.rowCountAttribute,
        props.columnCountAttribute, props.onTableChange, updateStats]);

    // ── Load from attribute ───────────────────────────────────────────────────
    useEffect(() => {
        if (isSaving) return;
        const raw = props.useAttributeData?.value || "";
        if (raw === lastSavedRef.current && lastSavedRef.current !== "") return;
        if (!raw) { if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500); return; }
        try {
            const td: TableData = JSON.parse(raw);
            if (!td.tableRows || td.rows <= 0 || td.columns <= 0) return;
            const validated = td.tableRows.map((row, ri) => ({
                ...row, id: `row_${ri + 1}`, rowIndex: ri + 1,
                cells: row.cells.map((cell, ci) => ({
                    id: `cell_${ri + 1}_${ci + 1}`,
                    sequenceNumber: cell.sequenceNumber ?? "-",
                    isBlocked:  cell.isBlocked  ?? false,
                    isMerged:   cell.isMerged   ?? false,
                    mergeId:    cell.mergeId    ?? "",
                    isBlank:    cell.isBlank    ?? false,
                    rowIndex:   ri + 1,
                    columnIndex: ci + 1,
                    checked:    cell.checked    ?? false,
                    isSelected: false,
                    rowSpan:    cell.rowSpan    ?? 1,
                    colSpan:    cell.colSpan    ?? 1,
                    isHidden:   cell.isHidden   ?? false
                } as CellObject))
            }));
            setRowCount(td.rows); setColumnCount(td.columns);
            ignoreAttrUpdateRef.current = true;
            if (props.rowCountAttribute?.status    === "available") props.rowCountAttribute.setValue(new Big(td.rows));
            if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(td.columns));
            setTableRows(validated);
            setSelectedCells(new Set());
            setDataLoaded(true);
            updateStats(validated);
            lastSavedRef.current = raw;
            if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
        } catch (e) {
            console.error("Table load error:", e);
            if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
        }
    }, [props.useAttributeData?.value, isSaving, isInitialLoad, updateStats,
        props.rowCountAttribute, props.columnCountAttribute]);

    useEffect(() => {
        if (ignoreAttrUpdateRef.current) { ignoreAttrUpdateRef.current = false; return; }
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value != null) {
            const v = Number(props.rowCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== rowCount && !isUserInputRef.current) setRowCount(v);
        }
    }, [props.rowCountAttribute?.value, rowCount]);

    useEffect(() => {
        if (ignoreAttrUpdateRef.current) { ignoreAttrUpdateRef.current = false; return; }
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value != null) {
            const v = Number(props.columnCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== columnCount && !isUserInputRef.current) setColumnCount(v);
        }
    }, [props.columnCountAttribute?.value, columnCount]);

    // ── Create table ──────────────────────────────────────────────────────────
    const createTable = useCallback((rc: number, cc: number) => {
        if (rc <= 0 || cc <= 0) return;
        const rows: TableRow[] = Array.from({ length: rc }, (_, ri) => ({
            id: `row_${ri + 1}`, rowIndex: ri + 1,
            cells: Array.from({ length: cc }, (_, ci) => ({
                id: `cell_${ri + 1}_${ci + 1}`,
                sequenceNumber: "-", isBlocked: false, isMerged: false, mergeId: "",
                isBlank: false, rowIndex: ri + 1, columnIndex: ci + 1,
                checked: false, isSelected: false, rowSpan: 1, colSpan: 1, isHidden: false
            } as CellObject))
        }));
        setTableRows(rows); setSelectedCells(new Set()); setDataLoaded(true);
        saveToBackend(rows, rc, cc);
    }, [saveToBackend]);

    useEffect(() => {
        const t = setTimeout(() => { if (!dataLoaded && tableRows.length === 0) createTable(rowCount, columnCount); }, 100);
        return () => clearTimeout(t);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [dataLoaded]);

    useEffect(() => { if (tableRows.length > 0) updateStats(tableRows); }, [tableRows, updateStats]);

    const applyDimensions = useCallback(() => {
        if (!rowCount || !columnCount || rowCount > 100 || columnCount > 100) { alert("Enter valid numbers (1-100)"); return; }
        ignoreAttrUpdateRef.current = true;
        if (props.rowCountAttribute?.status    === "available") props.rowCountAttribute.setValue(new Big(rowCount));
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(columnCount));
        createTable(rowCount, columnCount);
    }, [rowCount, columnCount, createTable, props.rowCountAttribute, props.columnCountAttribute]);

    const addRow = useCallback(() => {
        const n = rowCount + 1; if (n > 100) { alert("Max 100 rows"); return; }
        isUserInputRef.current = true; setRowCount(n);
        ignoreAttrUpdateRef.current = true;
        if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(n));
        setTableRows(prev => {
            const rows = [...prev, { id: `row_${n}`, rowIndex: n,
                cells: Array.from({ length: columnCount }, (_, ci) => ({
                    id: `cell_${n}_${ci + 1}`, sequenceNumber: "-", isBlocked: false, isMerged: false,
                    mergeId: "", isBlank: false, rowIndex: n, columnIndex: ci + 1,
                    checked: false, isSelected: false, rowSpan: 1, colSpan: 1, isHidden: false
                } as CellObject))
            }];
            saveToBackend(rows, n, columnCount); return rows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.rowCountAttribute, saveToBackend]);

    const addColumn = useCallback(() => {
        const n = columnCount + 1; if (n > 100) { alert("Max 100 columns"); return; }
        isUserInputRef.current = true; setColumnCount(n);
        ignoreAttrUpdateRef.current = true;
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(n));
        setTableRows(prev => {
            const rows = prev.map(row => ({ ...row, cells: [...row.cells, {
                id: `cell_${row.rowIndex}_${n}`, sequenceNumber: "-", isBlocked: false, isMerged: false,
                mergeId: "", isBlank: false, rowIndex: row.rowIndex, columnIndex: n,
                checked: false, isSelected: false, rowSpan: 1, colSpan: 1, isHidden: false
            } as CellObject]}));
            saveToBackend(rows, rowCount, n); return rows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.columnCountAttribute, saveToBackend]);

    // ── Cell value change ─────────────────────────────────────────────────────
    const handleCellValueChange = useCallback((rowIndex: number, colIndex: number, value: string) => {
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const cell = rows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
            if (!cell) return prev;
            cell.sequenceNumber = value;
            if (cell.mergeId) {
                const mid = cell.mergeId;
                rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === mid) c.sequenceNumber = value; }));
            }
            updateStats(rows);
            latestStateRef.current = { rows, rc: rowCount, cc: columnCount };
            if (props.autoSave) {
                if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
                saveTimerRef.current = setTimeout(() => {
                    if (latestStateRef.current) saveToBackend(latestStateRef.current.rows, latestStateRef.current.rc, latestStateRef.current.cc);
                }, 300);
            }
            return rows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [updateStats, props.autoSave, props.onCellClick, saveToBackend, rowCount, columnCount]);

    // ── Checkbox ──────────────────────────────────────────────────────────────
    const handleCheckboxChange = useCallback((rowIndex: number, colIndex: number) => {
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const cell = rows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
            if (!cell) return prev;
            const chk = !cell.checked;
            cell.checked = chk; cell.isBlocked = chk;
            if (cell.mergeId) {
                const mid = cell.mergeId;
                rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === mid) { c.checked = chk; c.isBlocked = chk; } }));
            }
            updateStats(rows);
            if (props.autoSave) saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [updateStats, props.autoSave, props.onCellClick, saveToBackend, rowCount, columnCount]);

    // ────────────────────────────────────────────────────────────────────────────
    // DRAG-SELECT SYSTEM
    //
    // Uses document-level mousemove/mouseup (NOT pointer capture).
    // mousedown on td/input → record drag start.
    // mousemove on document → elementFromPoint to find current cell → update rect.
    // mouseup on document   → if no movement = click/dblclick; if moved = drag end.
    //
    // Key: tableRowsRef always holds latest tableRows so handlers never go stale.
    // ────────────────────────────────────────────────────────────────────────────
    useEffect(() => {
        const onDocMouseMove = (e: MouseEvent) => {
            if (!dragging.current || !dragStart.current) return;
            const td = getTdCell(document.elementFromPoint(e.clientX, e.clientY));
            if (!td?.dataset.cellid) return;
            const pos = parseCellId(td.dataset.cellid);
            if (!pos) return;
            const s = dragStart.current;
            if (pos.row !== s.row || pos.col !== s.col) dragMoved.current = true;
            setSelectedCells(rectSet(s.row, s.col, pos.row, pos.col));
        };

        const onDocMouseUp = () => {
            if (!dragging.current) return;
            const moved = dragMoved.current;
            const start = dragStart.current;
            dragging.current  = false;
            dragStart.current = null;
            dragMoved.current = false;
            setIsDragging(false);

            if (!moved && start) {
                // Single or double click
                const cellId  = `cell_${start.row}_${start.col}`;
                const now     = Date.now();
                const isDbl   = (now - lastClickTime.current) < 350 && lastClickId.current === cellId;
                lastClickTime.current = now;
                lastClickId.current   = cellId;

                if (isDbl) {
                    // Double-click → select whole merge group (or toggle single cell)
                    const currentRows = tableRowsRef.current;
                    const clickedCell = currentRows
                        .find(r => r.rowIndex === start.row)
                        ?.cells.find(c => c.columnIndex === start.col);

                    setSelectedCells(prev => {
                        const next = new Set(prev);
                        if (clickedCell?.isMerged && clickedCell.mergeId) {
                            const ids = new Set<string>();
                            currentRows.forEach(r => r.cells.forEach(c => {
                                if (c.mergeId === clickedCell.mergeId) ids.add(c.id);
                            }));
                            const anySelected = Array.from(ids).some(id => next.has(id));
                            ids.forEach(id => anySelected ? next.delete(id) : next.add(id));
                        } else {
                            next.has(cellId) ? next.delete(cellId) : next.add(cellId);
                        }
                        return next;
                    });
                }
                // single click: selection was already set on mousedown (one cell)
            }
        };

        document.addEventListener("mousemove", onDocMouseMove);
        document.addEventListener("mouseup",   onDocMouseUp);
        return () => {
            document.removeEventListener("mousemove", onDocMouseMove);
            document.removeEventListener("mouseup",   onDocMouseUp);
        };
    }, []); // empty deps — uses only refs, never goes stale

    // Called from td onMouseDown
    const handleCellMouseDown = useCallback((rowIndex: number, colIndex: number, e: React.MouseEvent) => {
        // Checkboxes handle themselves — don't start drag
        if ((e.target as HTMLElement).tagName === "INPUT" && (e.target as HTMLInputElement).type === "checkbox") return;

        // For text inputs: do NOT preventDefault so the input keeps focus + caret.
        // For everything else: preventDefault stops browser text-selection during drag.
        if ((e.target as HTMLElement).tagName !== "INPUT") e.preventDefault();

        dragging.current  = true;
        dragStart.current = { row: rowIndex, col: colIndex };
        dragMoved.current = false;
        setIsDragging(true);
        // Immediately show the start cell as selected
        setSelectedCells(new Set([`cell_${rowIndex}_${colIndex}`]));
    }, []);

    // ── Merge ─────────────────────────────────────────────────────────────────
    const selectAll = useCallback(() => {
        const all = new Set<string>();
        tableRows.forEach(r => r.cells.forEach(c => { if (!c.isHidden) all.add(c.id); }));
        setSelectedCells(all);
    }, [tableRows]);

    const mergeCells = useCallback(() => {
        if (selectedCells.size < 2) return;
        const pos = Array.from(selectedCells).map(id => parseCellId(id)).filter(Boolean) as { row: number; col: number }[];
        const minR = Math.min(...pos.map(p => p.row)), maxR = Math.max(...pos.map(p => p.row));
        const minC = Math.min(...pos.map(p => p.col)), maxC = Math.max(...pos.map(p => p.col));
        if (selectedCells.size !== (maxR - minR + 1) * (maxC - minC + 1)) {
            alert("Please select a rectangular area to merge"); return;
        }
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            // Unmerge any existing merges in the selection area first
            for (let r = minR; r <= maxR; r++) for (let c = minC; c <= maxC; c++) {
                const cell = rows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                if (cell?.isMerged && cell.mergeId) {
                    const old = cell.mergeId;
                    rows.forEach(row => row.cells.forEach(cl => {
                        if (cl.mergeId === old) { cl.isMerged = false; cl.rowSpan = 1; cl.colSpan = 1; cl.isHidden = false; cl.mergeId = ""; }
                    }));
                }
            }
            const mid = `m_${minR}_${minC}_${maxR}_${maxC}`;
            const tl  = rows.find(r => r.rowIndex === minR)?.cells.find(c => c.columnIndex === minC);
            if (!tl) return prev;
            const val = tl.sequenceNumber, chk = tl.checked, blk = tl.isBlocked;
            for (let r = minR; r <= maxR; r++) for (let c = minC; c <= maxC; c++) {
                const cell = rows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                if (!cell) continue;
                cell.sequenceNumber = val; cell.checked = chk; cell.isBlocked = blk;
                cell.isMerged = true; cell.mergeId = mid;
                if (r === minR && c === minC) { cell.rowSpan = maxR - minR + 1; cell.colSpan = maxC - minC + 1; cell.isHidden = false; }
                else { cell.rowSpan = 1; cell.colSpan = 1; cell.isHidden = true; }
            }
            updateStats(rows); saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set([`cell_${minR}_${minC}`]));
    }, [selectedCells, updateStats, saveToBackend, rowCount, columnCount]);

    const unmergeCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const mids = new Set<string>();
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id); if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (cell?.isMerged && cell.mergeId) mids.add(cell.mergeId);
            });
            if (mids.size === 0) return prev;
            mids.forEach(mid => rows.forEach(r => r.cells.forEach(c => {
                if (c.mergeId === mid) { c.isMerged = false; c.rowSpan = 1; c.colSpan = 1; c.isHidden = false; c.mergeId = ""; }
            })));
            updateStats(rows); saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
    }, [selectedCells, updateStats, saveToBackend, rowCount, columnCount]);

    const blankCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id); if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (!cell) return;
                cell.isBlank = true;
                if (cell.mergeId) rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === cell.mergeId) c.isBlank = true; }));
            });
            updateStats(rows); saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateStats, saveToBackend, rowCount, columnCount]);

    const unblankCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id); if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (!cell) return;
                cell.isBlank = false;
                if (cell.mergeId) rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === cell.mergeId) c.isBlank = false; }));
            });
            updateStats(rows); saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateStats, saveToBackend, rowCount, columnCount]);

    // ── Unmount cleanup ───────────────────────────────────────────────────────
    useEffect(() => () => {
        if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
        if (latestStateRef.current && props.autoSave) {
            const { rows, rc, cc } = latestStateRef.current;
            const json = JSON.stringify({ rows: rc, columns: cc, tableRows: rows, metadata: { updatedAt: new Date().toISOString() } });
            if (props.useAttributeData?.status   === "available") props.useAttributeData.setValue(json);
            if (props.tableDataAttribute?.status === "available") props.tableDataAttribute.setValue(json);
        }
    }, [props.autoSave, props.useAttributeData, props.tableDataAttribute]);

    // ── Styles ────────────────────────────────────────────────────────────────
    const tblStyle      = { borderColor: props.tableBorderColor   || "#dee2e6" };
    const selStyle      = { backgroundColor: props.selectedCellColor || "#cfe2ff" };
    const mergeStyle    = { backgroundColor: props.mergedCellColor  || "#e3f2fd", borderColor: "#2196f3" };
    const blockedStyle  = { backgroundColor: "white", borderColor: "#fdd835" };
    const blankStyle    = { backgroundColor: "transparent", border: "none", borderColor: "transparent" };

    // ── Render ────────────────────────────────────────────────────────────────
    return (
        <div className={classNames("tableview-container", props.class)} style={props.style}>

            {/* Controls bar — visible only when needed */}
            {(props.showGenerateButton || (props.enableCellMerging && selectedCells.size > 0)) && (
                <div className="tableview-controls">
                    {props.showGenerateButton && (
                        <button className="tableview-btn tableview-btn-primary" onClick={applyDimensions}>
                            Generate Table
                        </button>
                    )}
                    {props.enableCellMerging && selectedCells.size > 0 && (
                        createElement("div", { style: { display: "contents" } },
                            createElement("div",    { className: "tableview-controls-divider" }),
                            createElement("p",      { className: "tableview-selection-info" },
                                `${selectedCells.size} cell(s) selected  •  click=toggle  •  drag=multi-select  •  dbl-click=group`),
                            createElement("button", { className: "tableview-btn tableview-btn-info",      onClick: selectAll },         "Select All"),
                            createElement("button", { className: "tableview-btn tableview-btn-warning",   onClick: mergeCells, disabled: selectedCells.size < 2 }, "Merge"),
                            createElement("button", { className: "tableview-btn tableview-btn-danger",    onClick: unmergeCells },       "Unmerge"),
                            props.enableBlankCells && createElement("button", { className: "tableview-btn tableview-btn-dark",    onClick: blankCells },   "Blank"),
                            props.enableBlankCells && createElement("button", { className: "tableview-btn tableview-btn-success", onClick: unblankCells }, "Unblank"),
                            createElement("button", { className: "tableview-btn tableview-btn-secondary", onClick: () => setSelectedCells(new Set()) }, "Clear")
                        )
                    )}
                </div>
            )}

            {/* Table section */}
            <div className="tableview-table-section">
                {props.showAddColumnButton && (
                    <div className="tableview-add-column-container">
                        <button className="tableview-btn tableview-btn-add" onClick={addColumn} title="Add Column">+</button>
                    </div>
                )}
                <div className="tableview-table-row-wrapper">
                    {props.showAddRowButton && (
                        <div className="tableview-add-row-container">
                            <button className="tableview-btn tableview-btn-add" onClick={addRow} title="Add Row">+</button>
                        </div>
                    )}

                    <div
                        className="tableview-table-wrapper"
                        style={{ userSelect: isDragging ? "none" : "auto" }}
                    >
                        <table className="tableview-table" style={tblStyle} data-rows={rowCount} data-cols={columnCount}>
                            <tbody>
                                {tableRows.map(row => (
                                    <tr key={row.id}>
                                        {row.cells.map(cell => {
                                            if (cell.isHidden) return null;

                                            // A merged cell is "selected" if any cell in its group is selected
                                            let isSel = selectedCells.has(cell.id);
                                            if (!isSel && cell.isMerged && cell.mergeId) {
                                                tableRows.forEach(r => r.cells.forEach(c => {
                                                    if (c.mergeId === cell.mergeId && selectedCells.has(c.id)) isSel = true;
                                                }));
                                            }

                                            return (
                                                <td
                                                    key={cell.id}
                                                    data-cellid={cell.id}
                                                    rowSpan={cell.rowSpan}
                                                    colSpan={cell.colSpan}
                                                    className={classNames("tableview-cell", {
                                                        "tableview-cell-merged":   cell.isMerged,
                                                        "tableview-cell-selected": isSel,
                                                        "tableview-cell-blocked":  cell.isBlocked && !cell.isBlank,
                                                        "tableview-cell-dragging": isDragging,
                                                        "tableview-cell-blank":    cell.isBlank
                                                    })}
                                                    onMouseDown={e => handleCellMouseDown(cell.rowIndex, cell.columnIndex, e)}
                                                    style={{
                                                        ...(cell.isBlank                    ? blankStyle   : {}),
                                                        ...(cell.isMerged && !cell.isBlank  ? mergeStyle   : {}),
                                                        ...(isSel                           ? selStyle     : {}),
                                                        ...(cell.isBlocked && !cell.isBlank ? blockedStyle : {})
                                                    }}
                                                >
                                                    <div
                                                        className="tableview-cell-content"
                                                        style={{ visibility: cell.isBlank ? "hidden" : "visible" }}
                                                    >
                                                        {/* Checkbox:
                                                            onMouseDown stops propagation so it never triggers
                                                            the td's drag handler */}
                                                        <input
                                                            type="checkbox"
                                                            className="tableview-checkbox"
                                                            checked={cell.checked}
                                                            disabled={!props.enableCheckbox}
                                                            onChange={() => handleCheckboxChange(cell.rowIndex, cell.columnIndex)}
                                                            onMouseDown={e => e.stopPropagation()}
                                                        />

                                                        {/* Text input:
                                                            - onChange updates the cell value normally.
                                                            - onMouseDown does NOT stopPropagation, so the event
                                                              bubbles to the <td> and starts the drag.
                                                            - We do NOT call preventDefault in the td handler
                                                              when target is an input, so the input gets focus
                                                              and the caret is placed correctly.
                                                            - This means typing, drag-to-select text inside the
                                                              input, AND drag-to-select cells all work together. */}
                                                        <input
                                                            type="text"
                                                            className="tableview-cell-input"
                                                            value={cell.sequenceNumber}
                                                            disabled={!props.enableCellEditing}
                                                            onChange={e => handleCellValueChange(cell.rowIndex, cell.columnIndex, e.target.value)}
                                                            placeholder="#"
                                                        />
                                                    </div>
                                                </td>
                                            );
                                        })}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            {/* Info bar */}
            <div className="tableview-info">
                <p><strong>Table:</strong>   {rowCount} × {columnCount} = {rowCount * columnCount} cells</p>
                <p><strong>Blocked:</strong>  {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isBlocked).length, 0)}</p>
                <p><strong>Merged:</strong>   {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isMerged && !c.isHidden).length, 0)}</p>
                <p><strong>Blank:</strong>    {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isBlank).length, 0)}</p>
            </div>
        </div>
    );
};

export default Tableview;
