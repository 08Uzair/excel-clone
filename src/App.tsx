import { useState, useEffect, useMemo, useRef } from "react";
import { useReactTable, getCoreRowModel } from "@tanstack/react-table";
import * as XLSX from "xlsx";

function getExcelColumnName(colIndex: number): string {
  let dividend = colIndex + 1;
  let columnName = "";
  let modulo;
  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return columnName;
}

type DropdownOption = {
  label: string;
  bgColor: string;
};

type DropdownsMap = {
  [cellKey: string]: DropdownOption[];
};

type SelectedDropdownValuesMap = {
  [cellKey: string]: string;
};

export default function App() {
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<any[]>([]);
  const [search, setSearch] = useState("");
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);

  const [selectedCell, setSelectedCell] = useState<{
    rowIndex: number;
    columnId: string;
  } | null>(null);

  const [selectedColumnId, setSelectedColumnId] = useState<string | null>(null);

  const [dropdowns, setDropdowns] = useState<DropdownsMap>({});
  const [selectedDropdownValues, setSelectedDropdownValues] =
    useState<SelectedDropdownValuesMap>({});

  const [showDropdownForm, setShowDropdownForm] = useState(false);
  const [newOptions, setNewOptions] = useState<DropdownOption[]>([
    { label: "", bgColor: "#ffffff" },
  ]);

  // --- Resizing State ---
  const [colWidths, setColWidths] = useState<{ [colId: string]: number }>({});
  const [rowHeights, setRowHeights] = useState<{ [rowIdx: number]: number }>(
    {}
  );
  const resizingCol = useRef<string | null>(null);
  const resizingRow = useRef<number | null>(null);
  const startX = useRef(0);
  const startY = useRef(0);
  const startWidth = useRef(0);
  const startHeight = useRef(0);

  // --- Resizing Logic ---
  useEffect(() => {
    const onMouseMove = (e: MouseEvent) => {
      if (resizingCol.current) {
        const delta = e.clientX - startX.current;
        setColWidths((old) => ({
          ...old,
          [resizingCol.current!]: Math.max(40, startWidth.current + delta),
        }));
      }
      if (resizingRow.current !== null) {
        const delta = e.clientY - startY.current;
        setRowHeights((old) => ({
          ...old,
          [resizingRow.current!]: Math.max(24, startHeight.current + delta),
        }));
      }
    };
    const onMouseUp = () => {
      resizingCol.current = null;
      resizingRow.current = null;
    };
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    return () => {
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
    };
  }, []);

  // --- Dropdown Modal Logic ---
  const updateOption = (
    idx: number,
    field: "label" | "bgColor",
    value: string
  ) => {
    setNewOptions((old) =>
      old.map((opt, i) => (i === idx ? { ...opt, [field]: value } : opt))
    );
  };
  const addOptionInput = () => {
    setNewOptions((old) => [...old, { label: "", bgColor: "#ffffff" }]);
  };
  const submitDropdown = () => {
    if (!selectedCell) return;
    const key = `${selectedCell.rowIndex}-${selectedCell.columnId}`;
    setDropdowns((old) => ({
      ...old,
      [key]: newOptions.filter((opt) => opt.label.trim() !== ""),
    }));
    setShowDropdownForm(false);
    setNewOptions([{ label: "", bgColor: "#ffffff" }]);
  };

  // --- Clipboard logic ---
  const serializeDropdownForClipboard = (key: string) => {
    if (!dropdowns[key]) return null;
    return JSON.stringify({
      __dropdown: true,
      options: dropdowns[key],
      selected: selectedDropdownValues[key] || "",
    });
  };

  const handleCopy = (e: React.ClipboardEvent, key: string) => {
    const serialized = serializeDropdownForClipboard(key);
    if (serialized) {
      e.clipboardData.setData("text/plain", serialized);
      e.preventDefault();
    }
  };

  const handlePaste = (
    e: React.ClipboardEvent,
    rowIndex: number,
    columnId: string
  ) => {
    const clipboardText = e.clipboardData.getData("text/plain");
    try {
      const data = JSON.parse(clipboardText);
      if (data.__dropdown) {
        const key = `${rowIndex}-${columnId}`;
        setDropdowns((old) => ({ ...old, [key]: data.options }));
        setSelectedDropdownValues((old) => ({ ...old, [key]: data.selected }));
        handleEdit(rowIndex, columnId, data.selected);
        e.preventDefault();
        return;
      }
    } catch {
      // Not JSON, normal paste
    }
  };

  const addRow = () => {
    if (columns.length === 0) return alert("Add a column first!");
    const emptyRow = columns.reduce((acc: any, col: any) => {
      acc[col.accessorKey] = "";
      return acc;
    }, {});
    setData((old) => [...old, emptyRow]);
  };

  const addColumn = () => {
    const newKey = getExcelColumnName(columns.length);
    setColumns((old) => [
      ...old,
      {
        accessorKey: newKey,
        header: newKey,
        cell: ({ row }: any) => {
          const key = `${row.index}-${newKey}`;
          const hasDropdown = dropdowns[key];
          if (hasDropdown) {
            const selectedLabel = selectedDropdownValues[key] || "";
            const option = dropdowns[key].find(
              (opt) => opt.label === selectedLabel
            );
            const bgColor = option ? option.bgColor : "transparent";

            return (
              <select
                className="border p-1 w-full"
                value={selectedLabel}
                onChange={(e) => {
                  const val = e.target.value;
                  setSelectedDropdownValues((old) => ({
                    ...old,
                    [key]: val,
                  }));
                  handleEdit(row.index, newKey, val);
                }}
                onFocus={() =>
                  setSelectedCell({ rowIndex: row.index, columnId: newKey })
                }
                onBlur={() => setSelectedCell(null)}
                style={{ backgroundColor: bgColor }}
                onCopy={(e) => handleCopy(e, key)}
                onPaste={(e) => handlePaste(e, row.index, newKey)}
              >
                <option value="" disabled>
                  Select...
                </option>
                {dropdowns[key].map((opt, idx) => (
                  <option
                    key={idx}
                    value={opt.label}
                    style={{ backgroundColor: opt.bgColor }}
                  >
                    {opt.label}
                  </option>
                ))}
              </select>
            );
          }
          return (
            <input
              value={row.original[newKey] || ""}
              onChange={(e) => handleEdit(row.index, newKey, e.target.value)}
              onFocus={() =>
                setSelectedCell({ rowIndex: row.index, columnId: newKey })
              }
              onBlur={() => setSelectedCell(null)}
              className="border p-1 w-full"
              style={{ backgroundColor: "transparent" }}
              onCopy={(e) => handleCopy(e, key)}
              onPaste={(e) => handlePaste(e, row.index, newKey)}
            />
          );
        },
      },
    ]);
    setData((old) => old.map((row) => ({ ...row, [newKey]: "" })));
  };

  const handleEdit = (rowIndex: number, key: string, value: string) => {
    setData((old) =>
      old.map((row, index) =>
        index === rowIndex ? { ...row, [key]: value } : row
      )
    );
    const cellKey = `${rowIndex}-${key}`;
    if (dropdowns[cellKey]) {
      setSelectedDropdownValues((old) => ({
        ...old,
        [cellKey]: value,
      }));
    }
  };

  const handleHeaderEdit = (accessorKey: string, value: string) => {
    setColumns((old) =>
      old.map((col) =>
        col.accessorKey === accessorKey ? { ...col, header: value } : col
      )
    );
  };

  const importSheet = (e: any) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const importedData = XLSX.utils.sheet_to_json<Record<string, any>>(ws);
      if (importedData.length > 0) {
        const importedColumns = Object.keys(importedData[0]).map(
          (key, index) => ({
            accessorKey: key,
            header: key,
            cell: ({ row }: any) => {
              const cellKey = `${row.index}-${key}`;
              const hasDropdown = dropdowns[cellKey];
              if (hasDropdown) {
                const selectedLabel = selectedDropdownValues[cellKey] || "";
                const option = dropdowns[cellKey].find(
                  (opt) => opt.label === selectedLabel
                );
                const bgColor = option ? option.bgColor : "transparent";

                return (
                  <select
                    className="border p-1 w-full"
                    value={selectedLabel}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedDropdownValues((old) => ({
                        ...old,
                        [cellKey]: val,
                      }));
                      handleEdit(row.index, key, val);
                    }}
                    onFocus={() =>
                      setSelectedCell({ rowIndex: row.index, columnId: key })
                    }
                    onBlur={() => setSelectedCell(null)}
                    style={{ backgroundColor: bgColor }}
                    onCopy={(e) => handleCopy(e, cellKey)}
                    onPaste={(e) => handlePaste(e, row.index, key)}
                  >
                    <option value="" disabled>
                      Select...
                    </option>
                    {dropdowns[cellKey].map((opt, idx) => (
                      <option
                        key={idx}
                        value={opt.label}
                        style={{ backgroundColor: opt.bgColor }}
                      >
                        {opt.label}
                      </option>
                    ))}
                  </select>
                );
              }
              return (
                <input
                  value={row.original[key] || ""}
                  onChange={(e) => handleEdit(row.index, key, e.target.value)}
                  onFocus={() =>
                    setSelectedCell({ rowIndex: row.index, columnId: key })
                  }
                  onBlur={() => setSelectedCell(null)}
                  className="border p-1 w-full"
                  style={{ backgroundColor: "transparent" }}
                  onCopy={(e) => handleCopy(e, cellKey)}
                  onPaste={(e) => handlePaste(e, row.index, key)}
                />
              );
            },
          })
        );
        setColumns(importedColumns);
        setData(importedData);
        setSelectedRowIndex(null);
      }
    };
    reader.readAsBinaryString(file);
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, "Sheet.xlsx");
  };

  const shareSheet = () => {
    navigator.clipboard.writeText(window.location.href);
    alert("Link copied to clipboard!");
  };

  const filteredData = useMemo(() => {
    if (!search) return data;
    return data.filter((row) =>
      Object.values(row).some((val) =>
        String(val).toLowerCase().includes(search.toLowerCase())
      )
    );
  }, [data, search]);

  const memoColumns = useMemo(() => columns, [columns]);

  const table = useReactTable({
    data: filteredData,
    columns: memoColumns,
    getCoreRowModel: getCoreRowModel(),
  });

  // Keyboard navigation and delete logic
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Delete") {
        // --- CELL DELETE ---
        if (selectedCell) {
          setData((old) =>
            old.map((row, idx) =>
              idx === selectedCell.rowIndex
                ? { ...row, [selectedCell.columnId]: "" }
                : row
            )
          );
          const cellKey = `${selectedCell.rowIndex}-${selectedCell.columnId}`;
          setDropdowns((old) => {
            const copy = { ...old };
            delete copy[cellKey];
            return copy;
          });
          setSelectedDropdownValues((old) => {
            const copy = { ...old };
            delete copy[cellKey];
            return copy;
          });
          e.preventDefault();
          return;
        }
        if (selectedColumnId) {
          setColumns((old) =>
            old.map((col) =>
              col.accessorKey === selectedColumnId
                ? { ...col, header: "" }
                : col
            )
          );
          setData((old) =>
            old.map((row) => ({
              ...row,
              [selectedColumnId]: "",
            }))
          );
          setSelectedDropdownValues((old) => {
            const updated = { ...old };
            Object.keys(updated).forEach((cellKey) => {
              if (cellKey.endsWith(`-${selectedColumnId}`)) {
                updated[cellKey] = "";
              }
            });
            return updated;
          });
          setDropdowns((old) => {
            const updated = { ...old };
            Object.keys(updated).forEach((cellKey) => {
              if (cellKey.endsWith(`-${selectedColumnId}`)) {
                delete updated[cellKey];
              }
            });
            return updated;
          });
          e.preventDefault();
          return;
        }
        if (
          selectedRowIndex !== null &&
          selectedRowIndex >= 0 &&
          selectedRowIndex < data.length
        ) {
          setData((old) => old.filter((_, i) => i !== selectedRowIndex));
          setSelectedRowIndex(null);
          e.preventDefault();
        }
      }
      if (["ArrowDown", "ArrowUp", "ArrowLeft", "ArrowRight"].includes(e.key)) {
        e.preventDefault();
        const inputs = document.querySelectorAll("input, select, textarea");
        const active = document.activeElement;
        const index = Array.from(inputs).indexOf(active as HTMLElement);
        if (index !== -1) {
          let nextIndex = index;
          const visibleColumnsCount = table.getVisibleFlatColumns().length;
          if (e.key === "ArrowRight") nextIndex = index + 1;
          if (e.key === "ArrowLeft") nextIndex = index - 1;
          if (e.key === "ArrowDown") nextIndex = index + visibleColumnsCount;
          if (e.key === "ArrowUp") nextIndex = index - visibleColumnsCount;
          if (inputs[nextIndex]) (inputs[nextIndex] as HTMLElement).focus();
        }
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [
    selectedRowIndex,
    data,
    table,
    selectedCell,
    selectedColumnId,
    dropdowns,
  ]);

  return (
    <div className="h-screen flex flex-col bg-white">
      <div className="flex justify-between items-center px-4 py-2 border-b bg-gray-50">
        <h1 className="font-semibold text-gray-800">React Spreadsheet</h1>
        <div className="flex gap-2 items-center">
          <input
            type="text"
            placeholder="Search..."
            className="border px-2 py-1 text-sm"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          <label className="border px-3 py-1 rounded hover:bg-gray-100 cursor-pointer">
            Import
            <input
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={importSheet}
            />
          </label>
          <button
            className="border px-3 py-1 rounded hover:bg-gray-100"
            onClick={exportToExcel}
          >
            Export
          </button>
          <button
            className="border px-3 py-1 rounded hover:bg-gray-100"
            onClick={shareSheet}
          >
            Share
          </button>
          <button
            className="border px-3 py-1 rounded hover:bg-gray-100"
            onClick={addColumn}
          >
            + Column
          </button>
          <button
            className="bg-green-600 text-white px-3 py-1 rounded hover:bg-green-700"
            onClick={addRow}
          >
            + Row
          </button>
          <button
            className="border px-3 py-1 rounded hover:bg-gray-100"
            onClick={() => {
              if (!selectedCell) {
                alert("Select a cell first!");
                return;
              }
              setShowDropdownForm(true);
            }}
          >
            Add Dropdown
          </button>
        </div>
      </div>

      {/* Dropdown options modal */}
      {showDropdownForm && (
        <div className="fixed inset-0 bg-black bg-opacity-30 flex items-center justify-center z-50">
          <div className="bg-white p-6 rounded shadow-lg w-96 max-h-[80vh] overflow-auto">
            <h2 className="text-lg font-semibold mb-4">Add Dropdown Options</h2>
            {newOptions.map((opt, idx) => (
              <div key={idx} className="flex items-center gap-2 mb-2">
                <input
                  type="text"
                  placeholder="Label"
                  value={opt.label}
                  onChange={(e) => updateOption(idx, "label", e.target.value)}
                  className="border p-1 flex-grow"
                />
                <input
                  type="color"
                  value={opt.bgColor}
                  onChange={(e) => updateOption(idx, "bgColor", e.target.value)}
                  className="w-12 h-8 p-0 border rounded"
                />
              </div>
            ))}
            <button
              onClick={addOptionInput}
              className="border px-3 py-1 rounded hover:bg-gray-100 mb-4"
            >
              + Add Option
            </button>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setShowDropdownForm(false)}
                className="border px-3 py-1 rounded hover:bg-gray-100"
              >
                Cancel
              </button>
              <button
                onClick={submitDropdown}
                className="bg-blue-600 text-white px-3 py-1 rounded hover:bg-blue-700"
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="flex-1 overflow-auto">
        <table
          className="min-w-full border-collapse text-sm w-full"
          style={{
            borderColor: "green",
            borderWidth: "1px",
            borderStyle: "solid",
            tableLayout: "fixed",
          }}
        >
          <thead className="sticky top-0 bg-white shadow">
            {/* Editable column names with resize handles */}
            <tr>
              <th
                style={{
                  borderColor: "green",
                  minWidth: 50,
                  width: 50,
                  background: "#f0f0f0",
                  position: "relative",
                }}
              ></th>
              {columns.map((col) => (
                <th
                  key={col.accessorKey}
                  style={{
                    borderColor: "green",
                    minWidth: 40,
                    width: colWidths[col.accessorKey] || 120,
                    position: "relative",
                    padding: "4px",
                  }}
                >
                  <input
                    type="text"
                    value={col.header as string}
                    onChange={(e) =>
                      handleHeaderEdit(col.accessorKey, e.target.value)
                    }
                    className="border p-1 w-full"
                    style={{
                      width: "100%",
                      boxSizing: "border-box",
                    }}
                    onFocus={() => setSelectedColumnId(col.accessorKey)}
                    onBlur={() => setSelectedColumnId(null)}
                  />
                  {/* --- Column resize handle --- */}
                  <div
                    style={{
                      position: "absolute",
                      right: 0,
                      top: 0,
                      width: 6,
                      height: "100%",
                      cursor: "col-resize",
                      zIndex: 10,
                      userSelect: "none",
                    }}
                    onMouseDown={(e) => {
                      resizingCol.current = col.accessorKey;
                      startX.current = e.clientX;
                      startWidth.current = colWidths[col.accessorKey] || 120;
                      e.preventDefault();
                    }}
                  />
                </th>
              ))}
            </tr>
            {/* Excel-style column letters */}
            <tr>
              <th
                style={{
                  borderColor: "green",
                  minWidth: 50,
                  width: 50,
                  background: "#ddd",
                  textAlign: "center",
                  fontWeight: "bold",
                }}
              >
                #
              </th>
              {columns.map((col, idx) => (
                <th
                  key={col.accessorKey + "-excel"}
                  style={{
                    borderColor: "green",
                    minWidth: 40,
                    width: colWidths[col.accessorKey] || 120,
                    background: "#ddd",
                    textAlign: "center",
                    fontWeight: "bold",
                    userSelect: "none",
                  }}
                >
                  {getExcelColumnName(idx)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {table.getRowModel().rows.map((row, rowIdx) => (
              <tr
                key={row.id}
                className="hover:bg-gray-50"
                style={{
                  borderColor: "green",
                  height: rowHeights[rowIdx] || 32,
                }}
              >
                {/* Row number with row resize handle */}
                <td
                  style={{
                    borderColor: "green",
                    minWidth: 50,
                    width: 50,
                    background: "#f9f9f9",
                    textAlign: "center",
                    userSelect: "none",
                    position: "relative",
                    height: rowHeights[rowIdx] || 32,
                  }}
                >
                  {rowIdx + 1}
                  {/* --- Row resize handle --- */}
                  <div
                    style={{
                      position: "absolute",
                      left: 0,
                      bottom: 0,
                      width: "100%",
                      height: 6,
                      cursor: "row-resize",
                      zIndex: 10,
                      userSelect: "none",
                    }}
                    onMouseDown={(e) => {
                      resizingRow.current = rowIdx;
                      startY.current = e.clientY;
                      startHeight.current = rowHeights[rowIdx] || 32;
                      e.preventDefault();
                    }}
                  />
                </td>
                {/* Data cells */}
                {row.getVisibleCells().map((cell) => {
                  const key = `${row.index}-${cell.column.id}`;
                  const isFocused =
                    selectedCell?.rowIndex === row.index &&
                    selectedCell?.columnId === cell.column.id;

                  const cellWidth = colWidths[cell.column.id] || 120;
                  const cellHeight = rowHeights[rowIdx] || 32;

                  const options = dropdowns[key];
                  if (options) {
                    const selectedLabel = selectedDropdownValues[key] || "";
                    const option = options.find(
                      (opt) => opt.label === selectedLabel
                    );
                    const bgColor = option ? option.bgColor : "transparent";

                    return (
                      <td
                        key={cell.id}
                        className="p-2 border select-none"
                        style={{
                          borderColor: isFocused ? "#064e03" : "green",
                          borderWidth: isFocused ? "3px" : "1px",
                          backgroundColor: bgColor,
                          transition: "border 0.3s",
                          width: cellWidth,
                          minWidth: 40,
                          maxWidth: 500,
                          height: cellHeight,
                          minHeight: 24,
                          position: "relative",
                          boxSizing: "border-box",
                        }}
                        onClick={() =>
                          setSelectedCell({
                            rowIndex: row.index,
                            columnId: cell.column.id,
                          })
                        }
                      >
                        <select
                          className="border-none p-1 w-full bg-transparent focus:outline-none"
                          value={selectedLabel}
                          onChange={(e) => {
                            const val = e.target.value;
                            setSelectedDropdownValues((old) => ({
                              ...old,
                              [key]: val,
                            }));
                            handleEdit(row.index, cell.column.id, val);
                          }}
                          onFocus={() =>
                            setSelectedCell({
                              rowIndex: row.index,
                              columnId: cell.column.id,
                            })
                          }
                          onBlur={() => setSelectedCell(null)}
                          style={{
                            backgroundColor: bgColor,
                            width: "100%",
                            height: "100%",
                          }}
                        >
                          <option value="" disabled>
                            Select...
                          </option>
                          {options.map((opt, idx) => (
                            <option
                              key={idx}
                              value={opt.label}
                              style={{ backgroundColor: opt.bgColor }}
                            >
                              {opt.label}
                            </option>
                          ))}
                        </select>
                      </td>
                    );
                  }
                  return (
                    <td
                      key={cell.id}
                      className="p-2 border"
                      style={{
                        borderColor: isFocused ? "#064e03" : "green",
                        borderWidth: isFocused ? "3px" : "1px",
                        backgroundColor: "transparent",
                        transition: "border 0.3s",
                        width: cellWidth,
                        minWidth: 40,
                        maxWidth: 500,
                        height: cellHeight,
                        minHeight: 24,
                        position: "relative",
                        boxSizing: "border-box",
                      }}
                      onClick={() =>
                        setSelectedCell({
                          rowIndex: row.index,
                          columnId: cell.column.id,
                        })
                      }
                    >
                      <input
                        value={cell.getValue() ?? ""}
                        onChange={(e) =>
                          handleEdit(row.index, cell.column.id, e.target.value)
                        }
                        onFocus={() =>
                          setSelectedCell({
                            rowIndex: row.index,
                            columnId: cell.column.id,
                          })
                        }
                        onBlur={() => setSelectedCell(null)}
                        className="border-none p-1 w-full bg-transparent focus:outline-none"
                        style={{
                          width: "100%",
                          height: "100%",
                        }}
                      />
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
