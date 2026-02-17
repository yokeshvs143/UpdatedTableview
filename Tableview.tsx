import { ReactElement, createElement, useState, useCallback, useEffect, useRef } from "react";
import classNames from "classnames";
import { TableviewContainerProps } from "../typings/TableviewProps";
import Big from "big.js";
import "./ui/Tableview.css";

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

// ─────────────────────────────────────────────────────────────────────────────
// Helper: walk up the DOM from any element and return the first <td data-cellid>
// ─────────────────────────────────────────────────────────────────────────────
function findCellTd(el: Element | null): HTMLElement | null {
    let cur = el;
    while (cur) {
        if (cur.tagName === "TD" && (cur as HTMLElement).dataset.cellid) return cur as HTMLElement;
        cur = cur.parentElement;
    }
    return null;
}

function parseCellId(cellId: string): { row: number; col: number } | null {
    // format: cell_ROW_COL
    const parts = cellId.replace("cell_", "").split("_");
    if (parts.length < 2) return null;
    return { row: parseInt(parts[0], 10), col: parseInt(parts[1], 10) };
}

function buildRectSet(r1: number, c1: number, r2: number, c2: number): Set<string> {
    const s = new Set<string>();
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++)
            s.add(`cell_${r}_${c}`);
    return s;
}

// ─────────────────────────────────────────────────────────────────────────────
const Tableview = (props: TableviewContainerProps): ReactElement => {

    const getInitialRows = () => {
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value)
            return Number(props.rowCountAttribute.value);
        return 3;
    };
    const getInitialColumns = () => {
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value)
            return Number(props.columnCountAttribute.value);
        return 3;
    };

    const [rowCount, setRowCount]           = useState<number>(getInitialRows());
    const [columnCount, setColumnCount]     = useState<number>(getInitialColumns());
    const [tableRows, setTableRows]         = useState<TableRow[]>([]);
    const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
    const [isDragging, setIsDragging]       = useState<boolean>(false);

    // ── Misc ─────────────────────────────────────────────────────────────────
    const [isInitialLoad, setIsInitialLoad] = useState<boolean>(true);
    const [isSaving, setIsSaving]           = useState<boolean>(false);
    const [dataLoaded, setDataLoaded]       = useState<boolean>(false);
    const lastSavedDataRef                  = useRef<string>("");
    const isUserInputRef                    = useRef<boolean>(false);
    const ignoreAttributeUpdateRef          = useRef<boolean>(false);
    const latestTableStateRef               = useRef<{ rows: TableRow[]; rowCount: number; columnCount: number } | null>(null);
    const saveTimeoutRef                    = useRef<NodeJS.Timeout | null>(null);

    // ── Pointer / drag refs (never cause re-render during drag) ───────────────
    const tableWrapperRef   = useRef<HTMLDivElement>(null);
    const isDraggingRef     = useRef<boolean>(false);
    const dragStartRef      = useRef<{ row: number; col: number } | null>(null);
    const dragHasMovedRef   = useRef<boolean>(false);
    // Track last double-click time to distinguish single vs double click
    const lastClickTimeRef  = useRef<number>(0);
    const lastClickCellRef  = useRef<string>("");

    // ── Statistics ────────────────────────────────────────────────────────────
    const updateCellStatistics = useCallback((rows: TableRow[]) => {
        const totalCells   = rows.reduce((s, r) => s + r.cells.length, 0);
        const blockedCells = rows.reduce((s, r) => s + r.cells.filter(c => c.isBlocked).length, 0);
        const mergedCells  = rows.reduce((s, r) => s + r.cells.filter(c => c.isMerged && !c.isHidden).length, 0);
        if (props.totalCellsAttribute?.status === "available")  props.totalCellsAttribute.setValue(new Big(totalCells));
        if (props.blockedCellsAttribute?.status === "available") props.blockedCellsAttribute.setValue(new Big(blockedCells));
        if (props.mergedCellsAttribute?.status === "available")  props.mergedCellsAttribute.setValue(new Big(mergedCells));
    }, [props.totalCellsAttribute, props.blockedCellsAttribute, props.mergedCellsAttribute]);

    // ── Load from attribute ───────────────────────────────────────────────────
    useEffect(() => {
        if (isSaving) return;
        const incomingData = props.useAttributeData?.value || "";
        if (incomingData === lastSavedDataRef.current && lastSavedDataRef.current !== "") return;
        if (incomingData) {
            try {
                const tableData: TableData = JSON.parse(incomingData);
                if (tableData.tableRows && tableData.rows > 0 && tableData.columns > 0) {
                    const validatedRows = tableData.tableRows.map((row, idx) => {
                        const rowIndex = idx + 1;
                        return {
                            ...row, id: `row_${rowIndex}`, rowIndex,
                            cells: row.cells.map((cell, cIdx) => {
                                const colIndex = cIdx + 1;
                                return {
                                    id: `cell_${rowIndex}_${colIndex}`,
                                    sequenceNumber: cell.sequenceNumber || "-",
                                    isBlocked:  cell.isBlocked  || false,
                                    isMerged:   cell.isMerged   || false,
                                    mergeId:    cell.mergeId    || "",
                                    isBlank:    cell.isBlank    || false,
                                    rowIndex,
                                    columnIndex: colIndex,
                                    checked:    cell.checked    || false,
                                    isSelected: false,
                                    rowSpan:    cell.rowSpan    || 1,
                                    colSpan:    cell.colSpan    || 1,
                                    isHidden:   cell.isHidden   || false
                                } as CellObject;
                            })
                        };
                    });
                    setRowCount(tableData.rows);
                    setColumnCount(tableData.columns);
                    ignoreAttributeUpdateRef.current = true;
                    if (props.rowCountAttribute?.status === "available")
                        props.rowCountAttribute.setValue(new Big(tableData.rows));
                    if (props.columnCountAttribute?.status === "available")
                        props.columnCountAttribute.setValue(new Big(tableData.columns));
                    setTableRows(validatedRows);
                    setSelectedCells(new Set());
                    setDataLoaded(true);
                    updateCellStatistics(validatedRows);
                    lastSavedDataRef.current = incomingData;
                    if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
                }
            } catch (e) {
                console.error("Error loading table:", e);
                if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
            }
        } else {
            if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
        }
    }, [props.useAttributeData?.value, updateCellStatistics, isSaving, isInitialLoad,
        props.rowCountAttribute, props.columnCountAttribute]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value != null) {
            const v = Number(props.rowCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== rowCount && !isUserInputRef.current) setRowCount(v);
        }
    }, [props.rowCountAttribute?.value, rowCount]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value != null) {
            const v = Number(props.columnCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== columnCount && !isUserInputRef.current) setColumnCount(v);
        }
    }, [props.columnCountAttribute?.value, columnCount]);

    // ── Save to backend ───────────────────────────────────────────────────────
    const saveToBackend = useCallback((rows: TableRow[], rowCnt: number, colCnt: number) => {
        const jsonData = JSON.stringify({
            rows: rowCnt, columns: colCnt, tableRows: rows,
            metadata: { updatedAt: new Date().toISOString() }
        });
        lastSavedDataRef.current = jsonData;
        setIsSaving(true);
        if (props.useAttributeData?.status === "available")   props.useAttributeData.setValue(jsonData);
        if (props.tableDataAttribute?.status === "available") props.tableDataAttribute.setValue(jsonData);
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status === "available")    props.rowCountAttribute.setValue(new Big(rowCnt));
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(colCnt));
        updateCellStatistics(rows);
        if (props.onTableChange?.canExecute) props.onTableChange.execute();
        setTimeout(() => setIsSaving(false), 200);
    }, [props.useAttributeData, props.tableDataAttribute, props.rowCountAttribute,
        props.columnCountAttribute, props.onTableChange, updateCellStatistics]);

    // ── Create new table ──────────────────────────────────────────────────────
    const createNewTable = useCallback((rows: number, cols: number) => {
        if (rows <= 0 || cols <= 0) return;
        const newTableRows: TableRow[] = Array.from({ length: rows }, (_, idx) => {
            const rowIndex = idx + 1;
            return {
                id: `row_${rowIndex}`, rowIndex,
                cells: Array.from({ length: cols }, (_, cIdx) => {
                    const colIndex = cIdx + 1;
                    return {
                        id: `cell_${rowIndex}_${colIndex}`,
                        sequenceNumber: "-", isBlocked: false, isMerged: false, mergeId: "",
                        isBlank: false, rowIndex, columnIndex: colIndex,
                        checked: false, isSelected: false, rowSpan: 1, colSpan: 1, isHidden: false
                    } as CellObject;
                })
            };
        });
        setTableRows(newTableRows);
        setSelectedCells(new Set());
        setDataLoaded(true);
        saveToBackend(newTableRows, rows, cols);
    }, [saveToBackend]);

    useEffect(() => {
        const t = setTimeout(() => {
            if (!dataLoaded && tableRows.length === 0) createNewTable(rowCount, columnCount);
        }, 100);
        return () => clearTimeout(t);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [dataLoaded]);

    useEffect(() => {
        if (tableRows.length > 0) updateCellStatistics(tableRows);
    }, [tableRows, updateCellStatistics]);

    // ── Dimension management ──────────────────────────────────────────────────
    const applyDimensions = useCallback(() => {
        if (isNaN(rowCount) || isNaN(columnCount) || rowCount <= 0 || columnCount <= 0) {
            alert("Please enter valid numbers"); return;
        }
        if (rowCount > 100 || columnCount > 100) { alert("Maximum 100 rows and 100 columns"); return; }
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status === "available")    props.rowCountAttribute.setValue(new Big(rowCount));
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(columnCount));
        createNewTable(rowCount, columnCount);
    }, [rowCount, columnCount, createNewTable, props.rowCountAttribute, props.columnCountAttribute]);

    const addRow = useCallback(() => {
        const n = rowCount + 1;
        if (n > 100) { alert("Maximum 100 rows"); return; }
        isUserInputRef.current = true;
        setRowCount(n);
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(n));
        setTableRows(prev => {
            const rows = [...prev];
            rows.push({
                id: `row_${n}`, rowIndex: n,
                cells: Array.from({ length: columnCount }, (_, cIdx) => ({
                    id: `cell_${n}_${cIdx + 1}`, sequenceNumber: "-", isBlocked: false,
                    isMerged: false, mergeId: "", isBlank: false, rowIndex: n,
                    columnIndex: cIdx + 1, checked: false, isSelected: false,
                    rowSpan: 1, colSpan: 1, isHidden: false
                } as CellObject))
            });
            saveToBackend(rows, n, columnCount);
            return rows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.rowCountAttribute, saveToBackend]);

    const addColumn = useCallback(() => {
        const n = columnCount + 1;
        if (n > 100) { alert("Maximum 100 columns"); return; }
        isUserInputRef.current = true;
        setColumnCount(n);
        ignoreAttributeUpdateRef.current = true;
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(n));
        setTableRows(prev => {
            const rows = prev.map(row => ({
                ...row,
                cells: [...row.cells, {
                    id: `cell_${row.rowIndex}_${n}`, sequenceNumber: "-", isBlocked: false,
                    isMerged: false, mergeId: "", isBlank: false, rowIndex: row.rowIndex,
                    columnIndex: n, checked: false, isSelected: false,
                    rowSpan: 1, colSpan: 1, isHidden: false
                } as CellObject]
            }));
            saveToBackend(rows, rowCount, n);
            return rows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.columnCountAttribute, saveToBackend]);

    // ── Cell value change ─────────────────────────────────────────────────────
    const handleCellValueChange = useCallback((rowIndex: number, colIndex: number, newValue: string) => {
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const cell = rows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
            if (!cell) return prev;
            cell.sequenceNumber = newValue;
            if (cell.mergeId) {
                const mid = cell.mergeId;
                rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === mid) c.sequenceNumber = newValue; }));
            }
            updateCellStatistics(rows);
            latestTableStateRef.current = { rows, rowCount, columnCount };
            if (props.autoSave) {
                if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
                saveTimeoutRef.current = setTimeout(() => {
                    if (latestTableStateRef.current) {
                        saveToBackend(latestTableStateRef.current.rows,
                            latestTableStateRef.current.rowCount,
                            latestTableStateRef.current.columnCount);
                    }
                }, 300);
            }
            return rows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ── Checkbox change ───────────────────────────────────────────────────────
    const handleCheckboxChange = useCallback((rowIndex: number, colIndex: number) => {
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const cell = rows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
            if (!cell) return prev;
            const checked = !cell.checked;
            cell.checked = checked; cell.isBlocked = checked;
            if (cell.mergeId) {
                const mid = cell.mergeId;
                rows.forEach(r => r.cells.forEach(c => {
                    if (c.mergeId === mid) { c.checked = checked; c.isBlocked = checked; }
                }));
            }
            updateCellStatistics(rows);
            if (props.autoSave) saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ─────────────────────────────────────────────────────────────────────────
    // POINTER-EVENT BASED DRAG SELECTION
    // Attached directly to the table wrapper div via useEffect so it works
    // identically regardless of what child element (input, checkbox, td) is
    // under the pointer — including after editing cell values.
    //
    // Strategy:
    //   pointerdown  → record start cell, call setPointerCapture so all future
    //                  pointer events come here even if mouse leaves the element
    //   pointermove  → update rectangular selection using elementFromPoint
    //   pointerup    → finalize; if mouse never moved it was a click → toggle
    //
    // We use native DOM events (not React synthetic) so we can call
    // setPointerCapture, which React doesn't expose cleanly.
    // ─────────────────────────────────────────────────────────────────────────
    useEffect(() => {
        const wrapper = tableWrapperRef.current;
        if (!wrapper) return;

        const getCellFromPoint = (x: number, y: number): { row: number; col: number; id: string } | null => {
            // Temporarily release pointer capture for elementFromPoint to work
            const el = document.elementFromPoint(x, y);
            const td = findCellTd(el);
            if (!td || !td.dataset.cellid) return null;
            const pos = parseCellId(td.dataset.cellid);
            if (!pos) return null;
            return { ...pos, id: td.dataset.cellid };
        };

        const onPointerDown = (e: PointerEvent) => {
            // Ignore right-click / middle-click
            if (e.button !== 0) return;

            // If the click target is a checkbox — let it do its own thing
            const target = e.target as HTMLElement;
            if (target.tagName === "INPUT" && (target as HTMLInputElement).type === "checkbox") return;

            const cell = getCellFromPoint(e.clientX, e.clientY);
            if (!cell) return;

            // Capture pointer so move/up fire here even if mouse leaves wrapper
            try { (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId); } catch (_) { /* ignore */ }

            isDraggingRef.current  = true;
            dragStartRef.current   = { row: cell.row, col: cell.col };
            dragHasMovedRef.current = false;

            // Don't prevent default for text inputs so they keep focus + caret
            if (!(target.tagName === "INPUT" && (target as HTMLInputElement).type === "text")) {
                e.preventDefault();
            }

            setIsDragging(true);
            // Start with just the clicked cell selected
            setSelectedCells(new Set([cell.id]));
        };

        const onPointerMove = (e: PointerEvent) => {
            if (!isDraggingRef.current || !dragStartRef.current) return;

            // When pointer is captured, elementFromPoint still returns the correct element
            // We need to release capture temporarily — instead use the captured coordinates
            // with a temporary releasePointerCapture trick, OR just compute from target chain.
            // The cleanest approach: use document.elementFromPoint directly.
            // With setPointerCapture, e.target is the capturing element but x/y are still correct.
            const cell = getCellFromPoint(e.clientX, e.clientY);
            if (!cell) return;

            const start = dragStartRef.current;
            if (cell.row !== start.row || cell.col !== start.col) {
                dragHasMovedRef.current = true;
            }

            setSelectedCells(buildRectSet(start.row, start.col, cell.row, cell.col));
        };

        const onPointerUp = (e: PointerEvent) => {
            if (!isDraggingRef.current) return;

            try { (e.currentTarget as HTMLElement).releasePointerCapture(e.pointerId); } catch (_) { /* ignore */ }

            const hasMoved    = dragHasMovedRef.current;
            const start       = dragStartRef.current;

            isDraggingRef.current  = false;
            dragStartRef.current   = null;
            dragHasMovedRef.current = false;
            setIsDragging(false);

            if (!hasMoved && start) {
                // It was a click (no drag movement) — handle single/double click
                const now      = Date.now();
                const cellId   = `cell_${start.row}_${start.col}`;
                const lastTime = lastClickTimeRef.current;
                const lastCell = lastClickCellRef.current;
                const isDouble = (now - lastTime) < 350 && lastCell === cellId;

                lastClickTimeRef.current = now;
                lastClickCellRef.current = cellId;

                if (isDouble) {
                    // Double click → select/deselect the entire merge group (or single cell)
                    handleDoubleClickSelect(start.row, start.col);
                } else {
                    // Single click → toggle this cell in the selection
                    setSelectedCells(prev => {
                        const next = new Set(prev);
                        if (next.has(cellId)) {
                            next.delete(cellId);
                        } else {
                            next.add(cellId);
                        }
                        return next;
                    });
                }

                if (props.onCellClick?.canExecute) props.onCellClick.execute();
            }
        };

        const onPointerCancel = () => {
            isDraggingRef.current   = false;
            dragStartRef.current    = null;
            dragHasMovedRef.current = false;
            setIsDragging(false);
        };

        wrapper.addEventListener("pointerdown",   onPointerDown);
        wrapper.addEventListener("pointermove",   onPointerMove);
        wrapper.addEventListener("pointerup",     onPointerUp);
        wrapper.addEventListener("pointercancel", onPointerCancel);

        return () => {
            wrapper.removeEventListener("pointerdown",   onPointerDown);
            wrapper.removeEventListener("pointermove",   onPointerMove);
            wrapper.removeEventListener("pointerup",     onPointerUp);
            wrapper.removeEventListener("pointercancel", onPointerCancel);
        };
        // tableRows is needed for handleDoubleClickSelect closure — but we
        // re-attach when tableRows changes via the dependency below
    }, [tableRows, props.onCellClick]); // eslint-disable-line react-hooks/exhaustive-deps

    // ── Double-click: select/deselect whole merge group ───────────────────────
    // Called from inside the pointer event handler above (closure over tableRows)
    const handleDoubleClickSelect = useCallback((rowIndex: number, colIndex: number) => {
        const cellId    = `cell_${rowIndex}_${colIndex}`;
        const clickedCell = tableRows
            .find(r => r.rowIndex === rowIndex)
            ?.cells.find(c => c.columnIndex === colIndex);

        setSelectedCells(prev => {
            const next = new Set(prev);

            if (clickedCell?.isMerged && clickedCell.mergeId) {
                // Collect all cell ids in this merge group
                const groupIds = new Set<string>();
                tableRows.forEach(r => r.cells.forEach(c => {
                    if (c.mergeId === clickedCell.mergeId) groupIds.add(c.id);
                }));
                const anySelected = Array.from(groupIds).some(id => next.has(id));
                groupIds.forEach(id => anySelected ? next.delete(id) : next.add(id));
            } else {
                // Single cell toggle
                next.has(cellId) ? next.delete(cellId) : next.add(cellId);
            }
            return next;
        });
    }, [tableRows]);

    // ── Cleanup on unmount ────────────────────────────────────────────────────
    useEffect(() => {
        return () => {
            if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
            if (latestTableStateRef.current && props.autoSave) {
                const { rows, rowCount: rc, columnCount: cc } = latestTableStateRef.current;
                const json = JSON.stringify({ rows: rc, columns: cc, tableRows: rows, metadata: { updatedAt: new Date().toISOString() } });
                if (props.useAttributeData?.status === "available")   props.useAttributeData.setValue(json);
                if (props.tableDataAttribute?.status === "available") props.tableDataAttribute.setValue(json);
            }
        };
    }, [props.autoSave, props.useAttributeData, props.tableDataAttribute]);

    // ── Merge / Unmerge / Blank / Unblank ────────────────────────────────────
    const createMergeId = (r1: number, c1: number, r2: number, c2: number) => `${r1}_${c1}_${r2}_${c2}`;

    const selectAllCells = useCallback(() => {
        const all = new Set<string>();
        tableRows.forEach(row => row.cells.forEach(c => { if (!c.isHidden) all.add(c.id); }));
        setSelectedCells(all);
    }, [tableRows]);

    const mergeCells = useCallback(() => {
        if (selectedCells.size < 2) return;
        const positions = Array.from(selectedCells).map(id => {
            const p = parseCellId(id);
            return p!;
        }).filter(Boolean);
        const minRow = Math.min(...positions.map(p => p.row));
        const maxRow = Math.max(...positions.map(p => p.row));
        const minCol = Math.min(...positions.map(p => p.col));
        const maxCol = Math.max(...positions.map(p => p.col));
        if (selectedCells.size !== (maxRow - minRow + 1) * (maxCol - minCol + 1)) {
            alert("Please select a rectangular area to merge"); return;
        }

        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            // Unmerge existing merges in area first
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = rows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                    if (cell?.isMerged && cell.mergeId) {
                        const oid = cell.mergeId;
                        rows.forEach(row => row.cells.forEach(cl => {
                            if (cl.mergeId === oid) {
                                cl.isMerged = false; cl.rowSpan = 1; cl.colSpan = 1;
                                cl.isHidden = false; cl.mergeId = "";
                            }
                        }));
                    }
                }
            }
            const mergeId   = createMergeId(minRow, minCol, maxRow, maxCol);
            const topLeft   = rows.find(r => r.rowIndex === minRow)?.cells.find(c => c.columnIndex === minCol);
            if (!topLeft) return prev;
            const val = topLeft.sequenceNumber, chk = topLeft.checked, blk = topLeft.isBlocked;
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = rows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                    if (!cell) continue;
                    cell.sequenceNumber = val; cell.checked = chk; cell.isBlocked = blk;
                    cell.isMerged = true; cell.mergeId = mergeId;
                    if (r === minRow && c === minCol) {
                        cell.rowSpan = maxRow - minRow + 1; cell.colSpan = maxCol - minCol + 1; cell.isHidden = false;
                    } else {
                        cell.rowSpan = 1; cell.colSpan = 1; cell.isHidden = true;
                    }
                }
            }
            updateCellStatistics(rows);
            saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set([`cell_${minRow}_${minCol}`]));
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const unmergeCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            const mergeIds = new Set<string>();
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id);
                if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (cell?.isMerged && cell.mergeId) mergeIds.add(cell.mergeId);
            });
            if (mergeIds.size === 0) return prev;
            mergeIds.forEach(mid => {
                rows.forEach(r => r.cells.forEach(c => {
                    if (c.mergeId === mid) {
                        c.isMerged = false; c.rowSpan = 1; c.colSpan = 1; c.isHidden = false; c.mergeId = "";
                    }
                }));
            });
            updateCellStatistics(rows);
            saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const blankCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id);
                if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (!cell) return;
                cell.isBlank = true;
                if (cell.mergeId) rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === cell.mergeId) c.isBlank = true; }));
            });
            updateCellStatistics(rows);
            saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const unblankCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prev => {
            const rows = prev.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) }));
            Array.from(selectedCells).forEach(id => {
                const p = parseCellId(id);
                if (!p) return;
                const cell = rows.find(r => r.rowIndex === p.row)?.cells.find(c => c.columnIndex === p.col);
                if (!cell) return;
                cell.isBlank = false;
                if (cell.mergeId) rows.forEach(r => r.cells.forEach(c => { if (c.mergeId === cell.mergeId) c.isBlank = false; }));
            });
            updateCellStatistics(rows);
            saveToBackend(rows, rowCount, columnCount);
            return rows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ── Styles ────────────────────────────────────────────────────────────────
    const tableStyle        = { borderColor: props.tableBorderColor   || "#dee2e6" };
    const selectedCellStyle = { backgroundColor: props.selectedCellColor || "#cfe2ff" };
    const mergedCellStyle   = { backgroundColor: props.mergedCellColor  || "#e3f2fd", borderColor: "#2196f3" };
    const blockedCellStyle  = { backgroundColor: "white", borderColor: "#fdd835" };
    const blankCellStyle    = { backgroundColor: "transparent", border: "none", borderColor: "transparent" };

    // ── Render ────────────────────────────────────────────────────────────────
    return (
        <div className={classNames("tableview-container", props.class)} style={props.style}>

            {/* Controls bar */}
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
                            createElement("p",      { className: "tableview-selection-info" }, `${selectedCells.size} cell(s) selected`),
                            createElement("button", { className: "tableview-btn tableview-btn-info",      onClick: selectAllCells, title: "Select all cells" }, "Select All"),
                            createElement("button", { className: "tableview-btn tableview-btn-warning",   onClick: mergeCells,     disabled: selectedCells.size < 2 }, "Merge Selected"),
                            createElement("button", { className: "tableview-btn tableview-btn-danger",    onClick: unmergeCells }, "Unmerge"),
                            props.enableBlankCells && createElement("button", { className: "tableview-btn tableview-btn-dark",    onClick: blankCells,   title: "Blank selected cells" }, "Blank"),
                            props.enableBlankCells && createElement("button", { className: "tableview-btn tableview-btn-success", onClick: unblankCells, title: "Unblank selected cells" }, "Unblank"),
                            createElement("button", { className: "tableview-btn tableview-btn-secondary", onClick: () => setSelectedCells(new Set()) }, "Clear Selection")
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

                    {/*
                     * tableWrapperRef: all pointer events are handled here via
                     * native addEventListener in the useEffect above.
                     * touch-action: none prevents browser scroll from fighting our drag.
                     * user-select: none during drag prevents text highlights.
                     */}
                    <div
                        ref={tableWrapperRef}
                        className="tableview-table-wrapper"
                        style={{
                            touchAction: isDragging ? "none" : "auto",
                            userSelect:  isDragging ? "none" : "auto"
                        }}
                    >
                        <table
                            className="tableview-table"
                            style={tableStyle}
                            data-rows={rowCount}
                            data-cols={columnCount}
                        >
                            <tbody>
                                {tableRows.map(row => (
                                    <tr key={row.id}>
                                        {row.cells.map(cell => {
                                            if (cell.isHidden) return null;

                                            let isSelected = selectedCells.has(cell.id);
                                            if (!isSelected && cell.isMerged && cell.mergeId) {
                                                tableRows.forEach(r => r.cells.forEach(c => {
                                                    if (c.mergeId === cell.mergeId && selectedCells.has(c.id))
                                                        isSelected = true;
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
                                                        "tableview-cell-selected": isSelected,
                                                        "tableview-cell-blocked":  cell.isBlocked && !cell.isBlank,
                                                        "tableview-cell-dragging": isDragging,
                                                        "tableview-cell-blank":    cell.isBlank
                                                    })}
                                                    style={{
                                                        ...(cell.isBlank              ? blankCellStyle    : {}),
                                                        ...(cell.isMerged && !cell.isBlank ? mergedCellStyle : {}),
                                                        ...(isSelected                ? selectedCellStyle : {}),
                                                        ...(cell.isBlocked && !cell.isBlank ? blockedCellStyle : {})
                                                    }}
                                                >
                                                    <div
                                                        className="tableview-cell-content"
                                                        style={{ visibility: cell.isBlank ? "hidden" : "visible" }}
                                                    >
                                                        {/* Checkbox — stopPropagation so pointer events
                                                            don't reach the wrapper and start a drag */}
                                                        <input
                                                            type="checkbox"
                                                            className="tableview-checkbox"
                                                            checked={cell.checked}
                                                            disabled={!props.enableCheckbox}
                                                            onChange={() => handleCheckboxChange(cell.rowIndex, cell.columnIndex)}
                                                            onPointerDown={e => e.stopPropagation()}
                                                        />

                                                        {/* Text input — onChange updates value.
                                                            NO stopPropagation on pointerDown so the
                                                            wrapper's pointerdown handler can start drag.
                                                            The wrapper skips preventDefault for inputs
                                                            so the input still gets focus + caret normally. */}
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
                <p><strong>Table:</strong> {rowCount} rows × {columnCount} columns = {rowCount * columnCount} cells</p>
                <p><strong>Blocked:</strong>  {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isBlocked).length, 0)}</p>
                <p><strong>Merged:</strong>   {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isMerged && !c.isHidden).length, 0)}</p>
                <p><strong>Blank:</strong>    {tableRows.reduce((s, r) => s + r.cells.filter(c => c.isBlank).length, 0)}</p>
            </div>
        </div>
    );
};

export default Tableview;
