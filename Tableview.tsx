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
    metadata?: { createdAt?: string; updatedAt?: string; };
}

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

    const [rowCount, setRowCount]       = useState<number>(getInitialRows());
    const [columnCount, setColumnCount] = useState<number>(getInitialColumns());
    const [tableRows, setTableRows]     = useState<TableRow[]>([]);
    const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());

    // ── Drag state ────────────────────────────────────────────────────────────
    const [isDragging, setIsDragging]           = useState<boolean>(false);
    const isDraggingRef      = useRef<boolean>(false);
    const dragStartCellRef   = useRef<{row:number,col:number}|null>(null);
    const dragCurrentCellRef = useRef<{row:number,col:number}|null>(null);
    const dragHasMovedRef    = useRef<boolean>(false);

    // ── Misc state ────────────────────────────────────────────────────────────
    const [isInitialLoad, setIsInitialLoad] = useState<boolean>(true);
    const [isSaving,      setIsSaving]      = useState<boolean>(false);
    const [dataLoaded,    setDataLoaded]    = useState<boolean>(false);
    const lastSavedDataRef       = useRef<string>("");
    const isUserInputRef         = useRef<boolean>(false);
    const ignoreAttributeUpdateRef = useRef<boolean>(false);
    const latestTableStateRef    = useRef<{rows:TableRow[],rowCount:number,columnCount:number}|null>(null);
    const saveTimeoutRef         = useRef<NodeJS.Timeout|null>(null);

    // ── Statistics ────────────────────────────────────────────────────────────
    const updateCellStatistics = useCallback((rows: TableRow[]) => {
        const totalCells   = rows.reduce((s,r)=>s+r.cells.length,0);
        const blockedCells = rows.reduce((s,r)=>s+r.cells.filter(c=>c.isBlocked).length,0);
        const mergedCells  = rows.reduce((s,r)=>s+r.cells.filter(c=>c.isMerged&&!c.isHidden).length,0);
        if (props.totalCellsAttribute?.status==="available")  props.totalCellsAttribute.setValue(new Big(totalCells));
        if (props.blockedCellsAttribute?.status==="available") props.blockedCellsAttribute.setValue(new Big(blockedCells));
        if (props.mergedCellsAttribute?.status==="available")  props.mergedCellsAttribute.setValue(new Big(mergedCells));
    }, [props.totalCellsAttribute, props.blockedCellsAttribute, props.mergedCellsAttribute]);

    // ── Load from attribute ───────────────────────────────────────────────────
    useEffect(() => {
        if (isSaving) return;
        const incomingData = props.useAttributeData?.value || "";
        if (incomingData === lastSavedDataRef.current && lastSavedDataRef.current !== "") return;

        if (incomingData && incomingData !== "") {
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
                                    rowIndex, columnIndex: colIndex,
                                    checked:   cell.checked   || false,
                                    isSelected: false,
                                    rowSpan:   cell.rowSpan   || 1,
                                    colSpan:   cell.colSpan   || 1,
                                    isHidden:  cell.isHidden  || false
                                } as CellObject;
                            })
                        };
                    });
                    setRowCount(tableData.rows);
                    setColumnCount(tableData.columns);
                    ignoreAttributeUpdateRef.current = true;
                    if (props.rowCountAttribute?.status==="available")    props.rowCountAttribute.setValue(new Big(tableData.rows));
                    if (props.columnCountAttribute?.status==="available") props.columnCountAttribute.setValue(new Big(tableData.columns));
                    setTableRows(validatedRows);
                    setSelectedCells(new Set());
                    setDataLoaded(true);
                    updateCellStatistics(validatedRows);
                    lastSavedDataRef.current = incomingData;
                    if (isInitialLoad) setTimeout(()=>setIsInitialLoad(false), 500);
                }
            } catch(e) {
                console.error("Error loading table:", e);
                if (isInitialLoad) setTimeout(()=>setIsInitialLoad(false), 500);
            }
        } else {
            if (isInitialLoad) setTimeout(()=>setIsInitialLoad(false), 500);
        }
    }, [props.useAttributeData?.value, updateCellStatistics, isSaving, isInitialLoad, props.rowCountAttribute, props.columnCountAttribute]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.rowCountAttribute?.status==="available" && props.rowCountAttribute.value != null) {
            const v = Number(props.rowCountAttribute.value);
            if (!isNaN(v) && v>0 && v<=100 && v!==rowCount && !isUserInputRef.current) setRowCount(v);
        }
    }, [props.rowCountAttribute?.value, rowCount]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.columnCountAttribute?.status==="available" && props.columnCountAttribute.value != null) {
            const v = Number(props.columnCountAttribute.value);
            if (!isNaN(v) && v>0 && v<=100 && v!==columnCount && !isUserInputRef.current) setColumnCount(v);
        }
    }, [props.columnCountAttribute?.value, columnCount]);

    // ── Save to backend ───────────────────────────────────────────────────────
    const saveToBackend = useCallback((rows: TableRow[], rowCnt: number, colCnt: number) => {
        const jsonData = JSON.stringify({ rows: rowCnt, columns: colCnt, tableRows: rows, metadata: { updatedAt: new Date().toISOString() } });
        lastSavedDataRef.current = jsonData;
        setIsSaving(true);
        if (props.useAttributeData?.status==="available")  props.useAttributeData.setValue(jsonData);
        if (props.tableDataAttribute?.status==="available") props.tableDataAttribute.setValue(jsonData);
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status==="available")    props.rowCountAttribute.setValue(new Big(rowCnt));
        if (props.columnCountAttribute?.status==="available") props.columnCountAttribute.setValue(new Big(colCnt));
        updateCellStatistics(rows);
        if (props.onTableChange?.canExecute) props.onTableChange.execute();
        setTimeout(()=>setIsSaving(false), 200);
    }, [props.useAttributeData, props.tableDataAttribute, props.rowCountAttribute, props.columnCountAttribute, props.onTableChange, updateCellStatistics]);

    // ── Create new table ──────────────────────────────────────────────────────
    const createNewTable = useCallback((rows: number, cols: number) => {
        if (rows<=0 || cols<=0) return;
        const newTableRows: TableRow[] = Array.from({length: rows}, (_,idx) => {
            const rowIndex = idx+1;
            return {
                id: `row_${rowIndex}`, rowIndex,
                cells: Array.from({length: cols}, (_,cIdx) => {
                    const colIndex = cIdx+1;
                    return { id:`cell_${rowIndex}_${colIndex}`, sequenceNumber:"-", isBlocked:false, isMerged:false, mergeId:"", isBlank:false, rowIndex, columnIndex:colIndex, checked:false, isSelected:false, rowSpan:1, colSpan:1, isHidden:false } as CellObject;
                })
            };
        });
        setTableRows(newTableRows);
        setSelectedCells(new Set());
        setDataLoaded(true);
        saveToBackend(newTableRows, rows, cols);
    }, [saveToBackend]);

    useEffect(() => {
        const timer = setTimeout(()=>{ if (!dataLoaded && tableRows.length===0) createNewTable(rowCount, columnCount); }, 100);
        return ()=>clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [dataLoaded]);

    useEffect(() => { if (tableRows.length>0) updateCellStatistics(tableRows); }, [tableRows, updateCellStatistics]);

    const applyDimensions = useCallback(() => {
        if (isNaN(rowCount)||isNaN(columnCount)||rowCount<=0||columnCount<=0) { alert("Please enter valid numbers"); return; }
        if (rowCount>100||columnCount>100) { alert("Maximum 100 rows and 100 columns"); return; }
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status==="available")    props.rowCountAttribute.setValue(new Big(rowCount));
        if (props.columnCountAttribute?.status==="available") props.columnCountAttribute.setValue(new Big(columnCount));
        createNewTable(rowCount, columnCount);
    }, [rowCount, columnCount, createNewTable, props.rowCountAttribute, props.columnCountAttribute]);

    const addRow = useCallback(() => {
        const newRowCount = rowCount+1;
        if (newRowCount>100) { alert("Maximum 100 rows"); return; }
        isUserInputRef.current = true;
        setRowCount(newRowCount);
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status==="available") props.rowCountAttribute.setValue(new Big(newRowCount));
        setTableRows(prevRows => {
            const newRows = [...prevRows];
            const rowIndex = newRowCount;
            newRows.push({ id:`row_${rowIndex}`, rowIndex, cells: Array.from({length:columnCount},(_,cIdx)=>({ id:`cell_${rowIndex}_${cIdx+1}`, sequenceNumber:"-", isBlocked:false, isMerged:false, mergeId:"", isBlank:false, rowIndex, columnIndex:cIdx+1, checked:false, isSelected:false, rowSpan:1, colSpan:1, isHidden:false } as CellObject)) });
            saveToBackend(newRows, newRowCount, columnCount);
            return newRows;
        });
        setTimeout(()=>{ isUserInputRef.current=false; }, 100);
    }, [rowCount, columnCount, props.rowCountAttribute, saveToBackend]);

    const addColumn = useCallback(() => {
        const newColCount = columnCount+1;
        if (newColCount>100) { alert("Maximum 100 columns"); return; }
        isUserInputRef.current = true;
        setColumnCount(newColCount);
        ignoreAttributeUpdateRef.current = true;
        if (props.columnCountAttribute?.status==="available") props.columnCountAttribute.setValue(new Big(newColCount));
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells:[...row.cells, { id:`cell_${row.rowIndex}_${newColCount}`, sequenceNumber:"-", isBlocked:false, isMerged:false, mergeId:"", isBlank:false, rowIndex:row.rowIndex, columnIndex:newColCount, checked:false, isSelected:false, rowSpan:1, colSpan:1, isHidden:false } as CellObject] }));
            saveToBackend(newRows, rowCount, newColCount);
            return newRows;
        });
        setTimeout(()=>{ isUserInputRef.current=false; }, 100);
    }, [rowCount, columnCount, props.columnCountAttribute, saveToBackend]);

    // ── Cell value change — works for normal AND merged cells ─────────────────
    // Key fix: when editing a merged cell, find ALL cells in the merge group
    // and update their sequenceNumber so the visible top-left cell is updated.
    const handleCellValueChange = useCallback((rowIndex: number, colIndex: number, newValue: string) => {
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells: row.cells.map(cell=>({...cell})) }));
            const targetCell = newRows.find(r=>r.rowIndex===rowIndex)?.cells.find(c=>c.columnIndex===colIndex);
            if (!targetCell) return prevRows;

            // Update the cell itself
            targetCell.sequenceNumber = newValue;

            // If this cell is part of a merge group, propagate value to all cells in the group
            // (including the visible top-left cell which may have a different rowIndex/colIndex)
            if (targetCell.mergeId && targetCell.mergeId !== "") {
                const mergeId = targetCell.mergeId;
                newRows.forEach(row => {
                    row.cells.forEach(cell => {
                        if (cell.mergeId === mergeId) {
                            cell.sequenceNumber = newValue;
                        }
                    });
                });
            }

            updateCellStatistics(newRows);
            latestTableStateRef.current = { rows: newRows, rowCount, columnCount };

            if (props.autoSave) {
                if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
                saveTimeoutRef.current = setTimeout(() => {
                    if (latestTableStateRef.current) {
                        saveToBackend(latestTableStateRef.current.rows, latestTableStateRef.current.rowCount, latestTableStateRef.current.columnCount);
                    }
                }, 300);
            }
            return newRows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ── Checkbox change ───────────────────────────────────────────────────────
    const handleCheckboxChange = useCallback((rowIndex: number, colIndex: number) => {
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells: row.cells.map(cell=>({...cell})) }));
            const targetCell = newRows.find(r=>r.rowIndex===rowIndex)?.cells.find(c=>c.columnIndex===colIndex);
            if (!targetCell) return prevRows;
            const newChecked = !targetCell.checked;
            targetCell.checked   = newChecked;
            targetCell.isBlocked = newChecked;
            if (targetCell.mergeId && targetCell.mergeId !== "") {
                const mergeId = targetCell.mergeId;
                newRows.forEach(row=>row.cells.forEach(cell=>{ if(cell.mergeId===mergeId){ cell.checked=newChecked; cell.isBlocked=newChecked; } }));
            }
            updateCellStatistics(newRows);
            if (props.autoSave) saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        if (props.onCellClick?.canExecute) props.onCellClick.execute();
    }, [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ── Rectangular selection helper ──────────────────────────────────────────
    const getRectangularSelection = useCallback((sr:number,sc:number,er:number,ec:number): Set<string> => {
        const sel = new Set<string>();
        for (let r=Math.min(sr,er); r<=Math.max(sr,er); r++)
            for (let c=Math.min(sc,ec); c<=Math.max(sc,ec); c++)
                sel.add(`cell_${r}_${c}`);
        return sel;
    }, []);

    // ── Drag: mousedown — record origin, do NOT select ────────────────────────
    const handleCellMouseDown = useCallback((rowIndex:number, colIndex:number, event:React.MouseEvent) => {
        const target = event.target as HTMLElement;
        const isCheckbox  = target.tagName==='INPUT' && (target as HTMLInputElement).type==='checkbox';
        const isTextInput = target.tagName==='INPUT' && (target as HTMLInputElement).type==='text';
        if (isCheckbox) return;
        if (!isTextInput) event.preventDefault(); // keep focus for text inputs
        isDraggingRef.current      = true;
        dragStartCellRef.current   = { row: rowIndex, col: colIndex };
        dragCurrentCellRef.current = { row: rowIndex, col: colIndex };
        dragHasMovedRef.current    = false;
        setIsDragging(true);
    }, []);

    // ── Drag: mouseenter — secondary update path ───────────────────────────────
    const handleCellMouseEnter = useCallback((rowIndex:number, colIndex:number) => {
        if (!isDraggingRef.current || !dragStartCellRef.current) return;
        const start = dragStartCellRef.current;
        if (start.row!==rowIndex || start.col!==colIndex) {
            dragHasMovedRef.current    = true;
            dragCurrentCellRef.current = { row:rowIndex, col:colIndex };
        }
        if (!dragHasMovedRef.current) return;
        setSelectedCells(getRectangularSelection(start.row, start.col, rowIndex, colIndex));
    }, [getRectangularSelection]);

    // ── Drag: global mousemove + mouseup ──────────────────────────────────────
    // mousemove uses elementFromPoint so it works over inputs, merged cells, blank cells.
    // mouseup always resets state — can never get stuck.
    useEffect(() => {
        const handleMouseMove = (e: MouseEvent) => {
            if (!isDraggingRef.current || !dragStartCellRef.current) return;
            let current: Element|null = document.elementFromPoint(e.clientX, e.clientY);
            while (current && !(current.tagName==='TD' && (current as HTMLElement).dataset.cellid))
                current = current.parentElement;
            if (!current) return;

            const parts = (current as HTMLElement).dataset.cellid!.replace('cell_','').split('_');
            const row = parseInt(parts[0]), col = parseInt(parts[1]);
            const start = dragStartCellRef.current;
            if (row!==start.row || col!==start.col) dragHasMovedRef.current = true;
            if (!dragHasMovedRef.current) return;

            dragCurrentCellRef.current = { row, col };
            const sel = new Set<string>();
            for (let r=Math.min(start.row,row); r<=Math.max(start.row,row); r++)
                for (let c=Math.min(start.col,col); c<=Math.max(start.col,col); c++)
                    sel.add(`cell_${r}_${c}`);
            setSelectedCells(sel);
        };

        const handleMouseUp = () => {
            // Always reset — prevents isDragging from ever getting stuck
            isDraggingRef.current      = false;
            dragStartCellRef.current   = null;
            dragCurrentCellRef.current = null;
            dragHasMovedRef.current    = false;
            setIsDragging(false);
        };

        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup',   handleMouseUp);
        return () => {
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup',   handleMouseUp);
        };
    }, []); // empty deps — refs never go stale

    // ── Cleanup on unmount ────────────────────────────────────────────────────
    useEffect(() => {
        return () => {
            if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
            if (latestTableStateRef.current && props.autoSave) {
                const { rows, rowCount:rc, columnCount:cc } = latestTableStateRef.current;
                const jsonData = JSON.stringify({ rows:rc, columns:cc, tableRows:rows, metadata:{ updatedAt:new Date().toISOString() } });
                if (props.useAttributeData?.status==="available")  props.useAttributeData.setValue(jsonData);
                if (props.tableDataAttribute?.status==="available") props.tableDataAttribute.setValue(jsonData);
            }
        };
    }, [props.autoSave, props.useAttributeData, props.tableDataAttribute]);

    // ── Double-click = toggle selection ──────────────────────────────────────
    const handleCellDoubleClick = useCallback((rowIndex:number, colIndex:number, event:React.MouseEvent) => {
        if (isDraggingRef.current) return;
        const target = event.target as HTMLElement;
        // Double-click on editable text input → select all text, not cell
        if (target.tagName==='INPUT' && (target as HTMLInputElement).type==='text' && props.enableCellEditing) return;
        if (target.tagName==='INPUT' && (target as HTMLInputElement).type==='checkbox') return;

        const cellId = `cell_${rowIndex}_${colIndex}`;
        if (props.onCellClick?.canExecute) props.onCellClick.execute();

        setSelectedCells(prev => {
            const newSet = new Set(prev);
            const clickedCell = tableRows.find(r=>r.rowIndex===rowIndex)?.cells.find(c=>c.columnIndex===colIndex);
            if (clickedCell?.isMerged && clickedCell.mergeId) {
                const ids = new Set<string>();
                tableRows.forEach(row=>row.cells.forEach(cell=>{ if(cell.mergeId===clickedCell.mergeId) ids.add(cell.id); }));
                const anySelected = Array.from(ids).some(id=>newSet.has(id));
                ids.forEach(id=>anySelected ? newSet.delete(id) : newSet.add(id));
            } else {
                newSet.has(cellId) ? newSet.delete(cellId) : newSet.add(cellId);
            }
            return newSet;
        });
    }, [props.onCellClick, tableRows, props.enableCellEditing]);

    // ── Merge / Unmerge / Blank / Unblank ────────────────────────────────────
    const createMergeId = (r1:number,c1:number,r2:number,c2:number) => `${r1}${c1}${r2}${c2}`;

    const selectAllCells = useCallback(() => {
        const all = new Set<string>();
        tableRows.forEach(row=>row.cells.forEach(cell=>{ if(!cell.isHidden) all.add(cell.id); }));
        setSelectedCells(all);
    }, [tableRows]);

    const mergeCells = useCallback(() => {
        if (selectedCells.size<2) return;
        const positions = Array.from(selectedCells).map(id=>{ const p=id.replace('cell_','').split('_'); return {row:parseInt(p[0]),col:parseInt(p[1])}; });
        const minRow=Math.min(...positions.map(p=>p.row)), maxRow=Math.max(...positions.map(p=>p.row));
        const minCol=Math.min(...positions.map(p=>p.col)), maxCol=Math.max(...positions.map(p=>p.col));
        if (selectedCells.size !== (maxRow-minRow+1)*(maxCol-minCol+1)) { alert("Please select a rectangular area to merge"); return; }

        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells:row.cells.map(cell=>({...cell})) }));
            // Unmerge existing merges in selection area
            for (let r=minRow; r<=maxRow; r++) for (let c=minCol; c<=maxCol; c++) {
                const cell = newRows.find(row=>row.rowIndex===r)?.cells.find(cl=>cl.columnIndex===c);
                if (cell?.isMerged && cell.mergeId) {
                    const oid = cell.mergeId;
                    newRows.forEach(row=>row.cells.forEach(cl=>{ if(cl.mergeId===oid){ cl.isMerged=false; cl.rowSpan=1; cl.colSpan=1; cl.isHidden=false; cl.mergeId=""; } }));
                }
            }
            const mergeId = createMergeId(minRow,minCol,maxRow,maxCol);
            const topLeft = newRows.find(r=>r.rowIndex===minRow)?.cells.find(c=>c.columnIndex===minCol);
            if (!topLeft) return prevRows;
            const mergedValue=topLeft.sequenceNumber, mergedChecked=topLeft.checked, mergedIsBlocked=topLeft.isBlocked;
            for (let r=minRow; r<=maxRow; r++) for (let c=minCol; c<=maxCol; c++) {
                const cell = newRows.find(row=>row.rowIndex===r)?.cells.find(cl=>cl.columnIndex===c);
                if (!cell) continue;
                cell.sequenceNumber=mergedValue; cell.checked=mergedChecked; cell.isBlocked=mergedIsBlocked;
                cell.isMerged=true; cell.mergeId=mergeId;
                if (r===minRow && c===minCol) { cell.rowSpan=maxRow-minRow+1; cell.colSpan=maxCol-minCol+1; cell.isHidden=false; }
                else { cell.rowSpan=1; cell.colSpan=1; cell.isHidden=true; }
            }
            updateCellStatistics(newRows);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        setSelectedCells(new Set([`cell_${minRow}_${minCol}`]));
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const unmergeCells = useCallback(() => {
        if (selectedCells.size===0) return;
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells:row.cells.map(cell=>({...cell})) }));
            const mergeIds = new Set<string>();
            Array.from(selectedCells).forEach(cellId=>{
                const p=cellId.replace('cell_','').split('_');
                const cell = newRows.find(r=>r.rowIndex===parseInt(p[0]))?.cells.find(c=>c.columnIndex===parseInt(p[1]));
                if (cell?.isMerged && cell.mergeId) mergeIds.add(cell.mergeId);
            });
            if (mergeIds.size===0) return prevRows;
            mergeIds.forEach(mergeId=>{
                newRows.forEach(row=>row.cells.forEach(cell=>{ if(cell.mergeId===mergeId){ cell.isMerged=false; cell.rowSpan=1; cell.colSpan=1; cell.isHidden=false; cell.mergeId=""; } }));
            });
            updateCellStatistics(newRows);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const blankCells = useCallback(() => {
        if (selectedCells.size===0) return;
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells:row.cells.map(cell=>({...cell})) }));
            Array.from(selectedCells).forEach(cellId=>{
                const p=cellId.replace('cell_','').split('_');
                const cell = newRows.find(r=>r.rowIndex===parseInt(p[0]))?.cells.find(c=>c.columnIndex===parseInt(p[1]));
                if (!cell) return;
                cell.isBlank=true;
                if (cell.mergeId) newRows.forEach(row=>row.cells.forEach(c=>{ if(c.mergeId===cell.mergeId) c.isBlank=true; }));
            });
            updateCellStatistics(newRows);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    const unblankCells = useCallback(() => {
        if (selectedCells.size===0) return;
        setTableRows(prevRows => {
            const newRows = prevRows.map(row=>({ ...row, cells:row.cells.map(cell=>({...cell})) }));
            Array.from(selectedCells).forEach(cellId=>{
                const p=cellId.replace('cell_','').split('_');
                const cell = newRows.find(r=>r.rowIndex===parseInt(p[0]))?.cells.find(c=>c.columnIndex===parseInt(p[1]));
                if (!cell) return;
                cell.isBlank=false;
                if (cell.mergeId) newRows.forEach(row=>row.cells.forEach(c=>{ if(c.mergeId===cell.mergeId) c.isBlank=false; }));
            });
            updateCellStatistics(newRows);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        setSelectedCells(new Set());
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount]);

    // ── Styles ────────────────────────────────────────────────────────────────
    const tableStyle       = { borderColor: props.tableBorderColor || '#dee2e6' };
    const selectedCellStyle = { backgroundColor: props.selectedCellColor || '#cfe2ff' };
    const mergedCellStyle  = { backgroundColor: props.mergedCellColor || '#e3f2fd', borderColor: '#2196f3' };
    const blockedCellStyle = { backgroundColor: 'white', borderColor: '#fdd835' };
    const blankCellStyle   = { backgroundColor: 'transparent', border: 'none', borderColor: 'transparent' };

    // ── Render ────────────────────────────────────────────────────────────────
    return (
        <div className={classNames("tableview-container", props.class)} style={props.style}>

            {/* Controls — only visible when needed */}
            {(props.showGenerateButton || (props.enableCellMerging && selectedCells.size > 0)) && (
                <div className="tableview-controls">
                    {props.showGenerateButton && (
                        <button className="tableview-btn tableview-btn-primary" onClick={applyDimensions}>
                            Generate Table
                        </button>
                    )}
                    {props.enableCellMerging && selectedCells.size > 0 && (
                        createElement('div', { style:{display:'contents'} },
                            createElement('div', { className:'tableview-controls-divider' }),
                            createElement('p', { className:'tableview-selection-info' }, `${selectedCells.size} cell(s) selected`),
                            createElement('button', { className:'tableview-btn tableview-btn-info',      onClick:selectAllCells,  title:'Select all cells' }, 'Select All'),
                            createElement('button', { className:'tableview-btn tableview-btn-warning',   onClick:mergeCells,      disabled:selectedCells.size<2 }, 'Merge Selected'),
                            createElement('button', { className:'tableview-btn tableview-btn-danger',    onClick:unmergeCells }, 'Unmerge'),
                            props.enableBlankCells && createElement('button', { className:'tableview-btn tableview-btn-dark',    onClick:blankCells,   title:'Blank selected cells' }, 'Blank'),
                            props.enableBlankCells && createElement('button', { className:'tableview-btn tableview-btn-success', onClick:unblankCells, title:'Unblank selected cells' }, 'Unblank'),
                            createElement('button', { className:'tableview-btn tableview-btn-secondary', onClick:()=>setSelectedCells(new Set()) }, 'Clear Selection')
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
                        style={{ userSelect: isDragging ? 'none' : 'auto' }}
                        onMouseUp={() => {
                            // Safety reset in case the document listener missed it
                            isDraggingRef.current=false; dragStartCellRef.current=null;
                            dragCurrentCellRef.current=null; dragHasMovedRef.current=false;
                            setIsDragging(false);
                        }}
                    >
                        <table className="tableview-table" style={tableStyle} data-rows={rowCount} data-cols={columnCount}>
                            <tbody>
                                {tableRows.map(row => (
                                    <tr key={row.id}>
                                        {row.cells.map(cell => {
                                            if (cell.isHidden) return null;

                                            // Determine selection — check merged group too
                                            let isSelected = selectedCells.has(cell.id);
                                            if (!isSelected && cell.isMerged && cell.mergeId) {
                                                tableRows.forEach(r=>r.cells.forEach(c=>{ if(c.mergeId===cell.mergeId && selectedCells.has(c.id)) isSelected=true; }));
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
                                                    onDoubleClick={(e) => handleCellDoubleClick(cell.rowIndex, cell.columnIndex, e)}
                                                    onMouseDown={(e)   => handleCellMouseDown(cell.rowIndex, cell.columnIndex, e)}
                                                    onMouseEnter={() => handleCellMouseEnter(cell.rowIndex, cell.columnIndex)}
                                                    style={{
                                                        ...(cell.isBlank  ? blankCellStyle   : {}),
                                                        ...(cell.isMerged && !cell.isBlank ? mergedCellStyle : {}),
                                                        ...(isSelected    ? selectedCellStyle : {}),
                                                        ...(cell.isBlocked && !cell.isBlank ? blockedCellStyle : {})
                                                    }}
                                                >
                                                    <div
                                                        className="tableview-cell-content"
                                                        style={{ visibility: cell.isBlank ? 'hidden' : 'visible' }}
                                                    >
                                                        {/* Checkbox — always visible, disabled when feature off */}
                                                        <input
                                                            type="checkbox"
                                                            className="tableview-checkbox"
                                                            checked={cell.checked}
                                                            disabled={!props.enableCheckbox}
                                                            onChange={(e) => { e.stopPropagation(); handleCheckboxChange(cell.rowIndex, cell.columnIndex); }}
                                                            onClick={(e) => e.stopPropagation()}
                                                            onMouseDown={(e) => e.stopPropagation()}
                                                        />
                                                        {/* Text input — always visible, disabled when feature off.
                                                            No onMouseDown stopPropagation: event bubbles to td
                                                            so handleCellMouseDown can start drag from here too. */}
                                                        <input
                                                            type="text"
                                                            className="tableview-cell-input"
                                                            value={cell.sequenceNumber}
                                                            disabled={!props.enableCellEditing}
                                                            onChange={(e) => handleCellValueChange(cell.rowIndex, cell.columnIndex, e.target.value)}
                                                            onClick={(e) => e.stopPropagation()}
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
                <p><strong>Table:</strong> {rowCount} rows × {columnCount} columns = {rowCount*columnCount} cells</p>
                <p><strong>Blocked Cells:</strong> {tableRows.reduce((s,r)=>s+r.cells.filter(c=>c.isBlocked).length,0)}</p>
                <p><strong>Merged Cells:</strong>  {tableRows.reduce((s,r)=>s+r.cells.filter(c=>c.isMerged&&!c.isHidden).length,0)}</p>
                <p><strong>Blank Cells:</strong>   {tableRows.reduce((s,r)=>s+r.cells.filter(c=>c.isBlank).length,0)}</p>
            </div>
        </div>
    );
};

export default Tableview;