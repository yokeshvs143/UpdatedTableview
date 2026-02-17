import { ReactElement, createElement } from "react";
import { TableviewPreviewProps } from "../typings/TableviewProps";

export function preview(props: TableviewPreviewProps): ReactElement {
    // Use default values since initialRows/initialColumns aren't in generated types
    const columns = 3;
    const rows = 3;

    return (
        <div className="tableview-preview">
            <div style={{
                padding: "10px",
                backgroundColor: "#f8f9fa",
                borderRadius: "8px",
                border: "2px dashed #dee2e6"
            }}>
                <div style={{
                    display: "flex",
                    gap: "10px",
                    marginBottom: "10px",
                    padding: "10px",
                    background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
                    borderRadius: "6px",
                    color: "white"
                }}>
                    <div style={{ fontSize: "12px", fontWeight: "600" }}>
                        📊 Tableview Widget Preview
                    </div>
                </div>
                
                <div style={{
                    backgroundColor: "white",
                    padding: "15px",
                    borderRadius: "6px",
                    boxShadow: "0 2px 4px rgba(0, 0, 0, 0.1)"
                }}>
                    <table style={{
                        borderCollapse: "collapse",
                        width: "100%"
                    }}>
                        <tbody>
                            {Array.from({ length: rows }, (_, rowIdx) => (
                                <tr key={rowIdx}>
                                    {Array.from({ length: columns }, (_, colIdx) => (
                                        <td key={colIdx} style={{
                                            border: "1px solid #dee2e6",
                                            padding: "8px",
                                            textAlign: "center",
                                            minWidth: "60px",
                                            height: "45px",
                                            backgroundColor: "#fff",
                                            fontSize: "11px"
                                        }}>
                                            {props.enableCheckbox && (
                                                <div style={{ marginBottom: "4px" }}>
                                                    <input type="checkbox" disabled style={{ width: "14px", height: "14px" }} />
                                                </div>
                                            )}
                                            {props.enableCellEditing && (
                                                <div style={{
                                                    border: "1px solid #ced4da",
                                                    borderRadius: "3px",
                                                    padding: "4px",
                                                    fontSize: "9px",
                                                    color: "#999"
                                                }}>
                                                    #
                                                </div>
                                            )}
                                        </td>
                                    ))}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
                
                <div style={{
                    marginTop: "10px",
                    padding: "8px",
                    backgroundColor: "white",
                    borderRadius: "6px",
                    fontSize: "11px",
                    color: "#495057"
                }}>
                    <strong>Configuration:</strong> {rows} rows × {columns} columns
                    {props.rowCountAttribute && <span> • Row attr: ✓</span>}
                    {props.columnCountAttribute && <span> • Col attr: ✓</span>}
                    {props.enableCellMerging && <span> • Merging: ✓</span>}
                </div>
            </div>
        </div>
    );
}

export function getPreviewCss(): string {
    return `
        .tableview-preview {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
        }
    `;
}