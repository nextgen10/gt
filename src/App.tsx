import React, { useState, useEffect } from 'react';
import {
  Plus,
  Trash2,
  Braces,
  AlertCircle,
  AlignLeft,
  WrapText,
  RotateCcw,
  Check,
  X,
  Download,
  Upload,
  Minus,
  Layout
} from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import yaml from 'js-yaml';

// --- Types ---
type Value = string | number | boolean | null | Value[] | { [key: string]: Value };

// --- Immutable Helpers ---
const setValue = (obj: any, path: (string | number)[], value: any): any => {
  if (path.length === 0) return value;
  const [head, ...tail] = path;
  if (Array.isArray(obj)) {
    const newArr = [...obj];
    newArr[head] = setValue(obj[head], tail, value);
    return newArr;
  }
  return { ...obj, [head]: setValue(obj[head] || {}, tail, value) };
};

const removeValue = (obj: any, path: (string | number)[]): any => {
  if (path.length === 0) return undefined;
  const [head, ...tail] = path;
  if (path.length === 1) {
    if (Array.isArray(obj)) return obj.filter((_, i) => i !== head);
    const newObj = { ...obj };
    delete newObj[head];
    return newObj;
  }
  if (Array.isArray(obj)) {
    const newArr = [...obj];
    newArr[head] = removeValue(obj[head], tail);
    return newArr;
  }
  return { ...obj, [head]: removeValue(obj[head], tail) };
};

// --- Components ---

const DynamicField = ({
  label,
  value,
  path,
  onChange,
  onRemove,
  onRename,
  hideHeader,
  level = 0,
  forceExpandState,
  highlightedPath,
  setHighlightedPath
}: {
  label: string | number,
  value: Value,
  path: (string | number)[],
  onChange: (path: (string | number)[], val: any) => void,
  onRemove?: (path: (string | number)[]) => void,
  onRename?: (newKey: string) => void,
  hideHeader?: boolean,
  level?: number,
  forceExpandState?: boolean | null,
  highlightedPath?: (string | number)[] | null,
  setHighlightedPath?: (path: (string | number)[]) => void
}) => {

  const isArray = Array.isArray(value);
  const isObject = value !== null && typeof value === 'object' && !isArray;
  const isPrimitive = !isArray && !isObject;

  const isTableMode = isArray && value.length > 0 &&
    value.every((item: any) => item !== null && typeof item === 'object' && !Array.isArray(item));

  const [isAdding, setIsAdding] = useState(false);
  const [newKey, setNewKey] = useState("");
  const [newType, setNewType] = useState("string");

  const [isExpanded, setIsExpanded] = useState(level === 0);

  useEffect(() => {
    if (forceExpandState !== null && forceExpandState !== undefined) {
      setIsExpanded(forceExpandState);
    }
  }, [forceExpandState]);

  const [isEditingKey, setIsEditingKey] = useState(false);
  const [editedKey, setEditedKey] = useState(String(label));

  useEffect(() => {
    setEditedKey(String(label));
  }, [label]);


  const handleAddArrayItem = () => {
    if (!isArray) return;
    let newItem: any = "";
    if (value.length > 0) {
      const cloneStructure = (item: any): any => {
        if (Array.isArray(item)) {
          if (item.length > 0) {
            return [cloneStructure(item[0])];
          }
          return [];
        }
        if (typeof item === 'object' && item !== null) {
          return Object.keys(item).reduce((acc, key) => ({ ...acc, [key]: cloneStructure(item[key]) }), {});
        }
        if (typeof item === 'number') return 0;
        if (typeof item === 'boolean') return false;
        return "";
      };
      newItem = cloneStructure(value[0]);
    }
    onChange(path, [...value, newItem]);
    if (setHighlightedPath) {
      setHighlightedPath([...path, value.length]);
      setIsExpanded(true);
    }
  };

  const handleAddField = () => {
    if (!newKey.trim()) return;

    let initialValue: any = "";
    switch (newType) {
      case "string": initialValue = ""; break;
      case "number": initialValue = 0; break;
      case "boolean": initialValue = false; break;
      case "array": initialValue = []; break;
      case "object": initialValue = {}; break;
    }

    onChange([...path, newKey], initialValue);
    if (setHighlightedPath) {
      setHighlightedPath([...path, newKey]);
      setIsExpanded(true);
    }
    setIsAdding(false);
    setNewKey("");
    setNewType("string");
  };


  const getTypeLabel = () => {
    if (isArray && isTableMode) return 'Table';
    if (isArray) return 'Array';
    if (isObject) return 'Object';
    return typeof value;
  };

  const isHighlighted = highlightedPath &&
    path.length === highlightedPath.length &&
    path.every((val, index) => val === highlightedPath[index]);

  return (
    <div
      className={`field-container ${isPrimitive ? 'input-group' : ''}`}
      style={isHighlighted ? {
        backgroundColor: 'rgba(34, 197, 94, 0.1)',
        border: '1px solid rgba(34, 197, 94, 0.4)',
        borderRadius: '6px',
        transition: 'all 0.5s ease'
      } : {}}
      ref={(el) => {
        if (isHighlighted && el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }}
    >
      {!hideHeader && (
        <div className="field-header">
          <label
            className="field-label select-none"
            style={{ cursor: isPrimitive ? 'default' : 'pointer', flex: 1, display: 'flex', alignItems: 'center' }}
            onClick={(e) => {
              if (!isPrimitive) {
                e.preventDefault();
                setIsExpanded(!isExpanded);
              }
            }}
          >
            {!isPrimitive && (
              <span
                className="text-blue-400 mr-2 hover:text-blue-300"
                onClick={(e) => { e.preventDefault(); setIsExpanded(!isExpanded); }}
              >
                {isExpanded ? <Minus size={14} /> : <Plus size={14} />}
              </span>
            )}

            {isEditingKey ? (
              <input
                autoFocus
                type="text"
                value={editedKey}
                onChange={(e) => setEditedKey(e.target.value)}
                onBlur={() => {
                  if (onRename) onRename(editedKey);
                  setIsEditingKey(false);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    if (onRename) onRename(editedKey);
                    setIsEditingKey(false);
                  }
                  if (e.key === 'Escape') {
                    setEditedKey(String(label));
                    setIsEditingKey(false);
                  }
                }}
                className="input py-0 px-1 h-6 text-sm w-32"
                onClick={(e) => e.stopPropagation()}
              />
            ) : (
              <span
                className="font-medium text-slate-200 hover:text-white cursor-text border-b border-transparent hover:border-slate-500 transition-colors"
                title="Click to rename key"
                onClick={(e) => {
                  if (onRename) {
                    e.preventDefault();
                    setIsEditingKey(true);
                  }
                }}
              >
                {label}
              </span>
            )}

            <span className="type-badge ml-2">{getTypeLabel()}</span>
          </label>

          {isArray && (
            <button
              onClick={(e) => { e.stopPropagation(); handleAddArrayItem(); }}
              className="btn btn-ghost hover:bg-slate-700 text-green-400 mr-1"
              title="Add Item"
              style={{ padding: '2px 6px' }}
            >
              <Plus size={14} />
            </button>
          )}

          {onRemove && (
            <button
              onClick={() => onRemove(path)}
              className="btn btn-danger-icon"
              title="Remove field"
            >
              <Trash2 size={14} />
            </button>
          )}
        </div>
      )}

      {hideHeader && !isPrimitive && (
        <div className="flex items-center gap-2 mb-1">
          <button
            onClick={(e) => { e.preventDefault(); setIsExpanded(!isExpanded); }}
            className="text-blue-400 hover:text-blue-300"
            title={isExpanded ? "Collapse" : "Expand"}
          >
            {isExpanded ? <Minus size={14} /> : <Plus size={14} />}
          </button>
          <span className="text-xs text-slate-500 font-medium select-none">{getTypeLabel()}</span>
        </div>
      )}

      {(isExpanded || isPrimitive) && (
        <>
          {isPrimitive && (
            <>
              {typeof value === 'boolean' ? (
                <div className="flex items-center gap-2">
                  <div
                    onClick={() => onChange(path, !value)}
                    className={`toggle-switch ${value ? 'active' : ''}`}
                  >
                    <div className="toggle-thumb" />
                  </div>
                  <span className="text-sm text-slate-300 ml-2">{value ? 'True' : 'False'}</span>
                </div>
              ) : typeof value === 'number' ? (
                <input
                  type="number"
                  value={value as number}
                  onChange={(e) => onChange(path, parseFloat(e.target.value) || 0)}
                  className="input input-number"
                />
              ) : (
                <input
                  type="text"
                  value={value as string}
                  onChange={(e) => onChange(path, e.target.value)}
                  className="input input-string"
                />
              )}
            </>
          )}

          {isObject && (
            <div className="nested-object">
              {Object.entries(value).map(([key, val]) => (
                <DynamicField
                  key={key}
                  label={key}
                  value={val}
                  path={[...path, key]}
                  onChange={onChange}
                  onRename={(newKeyName) => {
                    if (newKeyName === key) return;
                    if (!newKeyName.trim()) return;
                    if (Object.keys(value).includes(newKeyName)) {
                      alert("Key already exists!");
                      return;
                    }
                    const newObj = { ...value };
                    const keys = Object.keys(newObj);
                    const orderedObj: any = {};
                    keys.forEach((k) => {
                      if (k === key) {
                        orderedObj[newKeyName] = val;
                      } else {
                        orderedObj[k] = newObj[k];
                      }
                    });
                    onChange(path, orderedObj);
                  }}
                  onRemove={onRemove ? undefined : () => {
                    const newObj = { ...value };
                    delete newObj[key];
                    onChange(path, newObj);
                  }}
                  level={level + 1}
                  forceExpandState={forceExpandState}
                  highlightedPath={highlightedPath}
                  setHighlightedPath={setHighlightedPath}
                />
              ))}

              {isAdding ? (
                <div className="add-field-row">
                  <input
                    autoFocus
                    type="text"
                    placeholder="Key Name"
                    value={newKey}
                    onChange={(e) => setNewKey(e.target.value)}
                    className="input"
                    style={{ flex: 2 }}
                    onKeyDown={(e) => e.key === 'Enter' && handleAddField()}
                  />
                  <select
                    value={newType}
                    onChange={(e) => setNewType(e.target.value)}
                    className="select"
                    style={{ flex: 1 }}
                  >
                    <option value="string">String</option>
                    <option value="number">Number</option>
                    <option value="boolean">Boolean</option>
                    <option value="array">Array</option>
                    <option value="object">Object</option>
                  </select>
                  <button onClick={handleAddField} className="btn btn-primary" title="Confirm">
                    <Check size={16} />
                  </button>
                  <button onClick={() => setIsAdding(false)} className="btn btn-danger-icon" title="Cancel">
                    <X size={16} />
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => setIsAdding(true)}
                  className="btn btn-ghost"
                  style={{ marginTop: '0.5rem', marginLeft: '0.5rem', fontSize: '0.75rem' }}
                >
                  <Plus size={14} /> Add Field
                </button>
              )}
            </div>
          )}

          {isArray && (
            isTableMode ? (
              <div className="json-table-container">
                <table className="json-table">
                  <thead>
                    <tr>
                      {Array.from(new Set(value.flatMap((item: any) => Object.keys(item)))).map(key => {
                        const sampleVal = value.find((v: any) => v && v[key] !== undefined && v[key] !== null) as any;
                        const typeOfKey = sampleVal ? typeof sampleVal[key] : 'any';
                        return <th key={key}>{key} <span style={{ opacity: 0.5, fontSize: '0.8em', fontWeight: 'normal' }}>({typeOfKey})</span></th>
                      })}
                      <th style={{ width: '40px' }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {value.map((item: any, rowIndex: number) => {
                      const allKeys = Array.from(new Set(value.flatMap((it: any) => Object.keys(it))));
                      return (
                        <tr key={rowIndex}>
                          {allKeys.map(key => (
                            <td key={key}>
                              <DynamicField
                                label={key}
                                value={item[key] === undefined ? null : item[key]}
                                path={[...path, rowIndex, key]}
                                onChange={onChange}
                                hideHeader={true}
                                level={level + 1}
                                forceExpandState={forceExpandState}
                                highlightedPath={highlightedPath}
                                setHighlightedPath={setHighlightedPath}
                              />
                            </td>
                          ))}
                          <td>
                            <button
                              onClick={() => onChange(path, value.filter((_: any, i: number) => i !== rowIndex))}
                              className="btn btn-danger-icon p-1"
                              title="Remove Row"
                            >
                              <Trash2 size={14} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <button onClick={handleAddArrayItem} className="btn btn-primary" style={{ width: '100%', marginTop: '0.5rem', justifyContent: 'center' }}>
                  <Plus size={16} /> Add Row
                </button>
              </div>
            ) : (
              <div className="nested-object">
                {value.map((item, index) => (
                  <div key={index} className="array-item">
                    <DynamicField
                      label={`Item ${index + 1}`}
                      value={item}
                      path={[...path, index]}
                      onChange={onChange}
                      onRemove={(p) => onRemove ? onRemove(p) : onChange(path, value.filter((_, i) => i !== index))}
                      level={level + 1}
                      forceExpandState={forceExpandState}
                    />
                  </div>
                ))}
                <button onClick={handleAddArrayItem} className="btn btn-primary" style={{ width: '100%', marginTop: '0.5rem', justifyContent: 'center' }}>
                  <Plus size={16} /> Add Item
                </button>
              </div>
            )
          )}
        </>
      )}

      {!isExpanded && !isPrimitive && (
        <div
          className="px-4 py-2 text-xs text-slate-500 italic border-l-2 border-slate-700 ml-2 cursor-pointer hover:text-slate-300 transition-colors"
          onClick={() => setIsExpanded(true)}
          title="Click to expand"
        >
          ... {isArray ? `${value.length} items` : `${Object.keys(value).length} keys`} collapsed
        </div>
      )}
    </div>
  );
};


const RootObjectAdder = ({ onChange, rootData, setHighlightedPath }: { onChange: any, rootData: any, setHighlightedPath: any }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [newKey, setNewKey] = useState("");
  const [newType, setNewType] = useState("string");

  const handleAddField = () => {
    if (!newKey.trim()) return;

    if (rootData && typeof rootData === 'object' && !Array.isArray(rootData) && newKey in rootData) {
      alert("Key already exists!");
      return;
    }

    let initialValue: any = "";
    switch (newType) {
      case "string": initialValue = ""; break;
      case "number": initialValue = 0; break;
      case "boolean": initialValue = false; break;
      case "array": initialValue = []; break;
      case "object": initialValue = {}; break;
    }
    onChange([newKey], initialValue);
    if (setHighlightedPath) {
      setHighlightedPath([newKey]);
    }
    setIsAdding(false);
    setNewKey("");
  };

  if (isAdding) {
    return (
      <div className="add-field-row">
        <input
          autoFocus
          type="text"
          placeholder="Key Name"
          value={newKey}
          onChange={(e) => setNewKey(e.target.value)}
          className="input"
          style={{ flex: 2 }}
          onKeyDown={(e) => e.key === 'Enter' && handleAddField()}
        />
        <select
          value={newType}
          onChange={(e) => setNewType(e.target.value)}
          className="select"
          style={{ flex: 1 }}
        >
          <option value="string">String</option>
          <option value="number">Number</option>
          <option value="boolean">Boolean</option>
          <option value="array">Array</option>
          <option value="object">Object</option>
        </select>
        <button onClick={handleAddField} className="btn btn-primary" title="Confirm">
          <Check size={16} />
        </button>
        <button onClick={() => setIsAdding(false)} className="btn btn-danger-icon" title="Cancel">
          <X size={16} />
        </button>
      </div>
    )
  }

  return (
    <button
      onClick={() => setIsAdding(true)}
      className="btn btn-ghost"
      style={{ marginTop: '1rem', width: '100%', justifyContent: 'center' }}
    >
      <Plus size={14} /> Add Field to Root
    </button>
  );
};

// --- Main App ---

function App() {
  const [jsonInput, setJsonInput] = useState<string>('{\n  "title": "My Awesome Form",\n  "isActive": true,\n  "count": 42,\n  "tags": ["hero", "dark-mode", "react"],\n  "author": {\n    "name": "Jane Doe",\n    "email": "jane@example.com"\n  },\n  "features": [\n    { "id": 1, "name": "Login" },\n    { "id": 2, "name": "Dashboard" }\n  ]\n}');
  const [parsedData, setParsedData] = useState<Value | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [wordWrap, setWordWrap] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<'json' | 'form' | 'yaml'>('json');
  const [yamlInput, setYamlInput] = useState<string>('');
  const [forceExpandState, setForceExpandState] = useState<boolean | null>(null);
  const [highlightedPath, setHighlightedPath] = useState<(string | number)[] | null>(null);

  useEffect(() => {
    try {
      const parsed = JSON.parse(jsonInput);
      setParsedData(parsed);
      setError(null);
    } catch (e) {
      if (e instanceof Error) setError(e.message);
    }
  }, []);

  const handleJsonInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newVal = e.target.value;
    setJsonInput(newVal);
    try {
      const parsed = JSON.parse(newVal);
      setParsedData(parsed);
      setError(null);
      setForceExpandState(false);
    } catch (e) {
      if (e instanceof Error) setError(e.message);
    }
  };

  const handleFormChange = (path: (string | number)[], newValue: any) => {
    if (!parsedData) return;
    const newData = setValue(parsedData, path, newValue);
    setParsedData(newData);
    setJsonInput(JSON.stringify(newData, null, 2));
  };

  const handleFormRemove = (path: (string | number)[]) => {
    if (!parsedData) return;
    const newData = removeValue(parsedData, path);
    setParsedData(newData);
    setJsonInput(JSON.stringify(newData, null, 2));
  };

  const formatJson = () => {
    try {
      const parsed = JSON.parse(jsonInput);
      setJsonInput(JSON.stringify(parsed, null, 2));
      setError(null);
    } catch (e) { /* ignore */ }
  };

  const handleExportExcel = async () => {
    if (!parsedData) return;

    type ExcelRow = {
      type: 'field' | 'table-header' | 'table-col-headers' | 'table-row';
      path: string;
      label: string;
      value?: any;
      headers?: string[];
      values?: any[];
      level: number;
    };

    const rows: ExcelRow[] = [];

    const traverse = (obj: any, path: string, key: string, level: number) => {
      const isTable = Array.isArray(obj) && obj.length > 0 &&
        obj.every((i: any) => i && typeof i === 'object' && !Array.isArray(i));

      if (isTable) {
        rows.push({ type: 'table-header', path: path, label: key, level });

        const keys = new Set<string>();
        obj.forEach((row: any) => Object.keys(row).forEach(k => keys.add(k)));
        const colHeaders = Array.from(keys);

        rows.push({ type: 'table-col-headers', path: '', label: '', headers: colHeaders, level: level + 1 });

        obj.forEach((rowObj: any, rowIndex: number) => {
          const rowVals = colHeaders.map(k => {
            const val = rowObj[k];
            if (val && typeof val === 'object') {
              return Array.isArray(val) ? `[Array(${val.length})]` : '[Object]';
            }
            return val;
          });
          rows.push({ type: 'table-row', path: '', label: '', values: rowVals, level: level + 1 });

          Object.entries(rowObj).forEach(([k, v]) => {
            if (v && typeof v === 'object') {
              const childPath = path ? `${path}.${rowIndex}.${k}` : `${rowIndex}.${k}`;
              traverse(v, childPath, k, level + 2);
            }
          });
        });
        return;
      }

      if (Array.isArray(obj)) {
        obj.forEach((v: any, i: number) => {
          traverse(v, path ? `${path}.${i}` : `${i}`, `Item ${i + 1}`, level);
        });
        return;
      }

      if (obj && typeof obj === 'object') {
        if (key && key !== 'Root') {
          rows.push({ type: 'table-header', path: path, label: key, level });
          level++;
        }

        Object.entries(obj).forEach(([k, v]) => {
          traverse(v, path ? `${path}.${k}` : k, k, level);
        });
        return;
      }

      rows.push({ type: 'field', path: path, label: key, value: obj, level });
    };

    traverse(parsedData, '', '', 0);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data", { views: [{ showGridLines: false }] });

    const mainHeader = worksheet.addRow(['Structure', 'Value']);
    mainHeader.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    mainHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
    mainHeader.height = 30;

    worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };

    let tableStripingIndex = 0;

    rows.forEach(r => {
      let row: ExcelJS.Row;
      const indentStr = "  ".repeat(r.level);

      if (r.type === 'field') {
        row = worksheet.addRow([indentStr + r.label, r.value]);

        const pathCell = row.getCell(1);
        pathCell.font = { color: { argb: 'FF374151' }, bold: true, name: 'Calibri', size: 11 };
        pathCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        row.height = 22;

        const valCell = row.getCell(2);
        valCell.alignment = { vertical: 'middle', wrapText: true, indent: 1 };

        // Subtle Separator
        row.eachCell(c => c.border = { bottom: { style: 'dotted', color: { argb: 'FFCBD5E1' } } });

        if (typeof r.value === 'boolean') {
          valCell.value = r.value ? 'TRUE' : 'FALSE';
          valCell.font = { color: { argb: 'FF7C3AED' }, bold: true };
          valCell.dataValidation = { type: 'list', allowBlank: false, formulae: ['"TRUE,FALSE"'] };
        } else if (typeof r.value === 'number') {
          valCell.font = { color: { argb: 'FF0369A1' }, bold: true };
          valCell.alignment = { horizontal: 'left', indent: 1 };
        } else {
          valCell.font = { color: { argb: 'FF1F2937' } };
        }
      }

      else if (r.type === 'table-header') {
        row = worksheet.addRow([indentStr + r.label]);
        const cell = row.getCell(1);
        cell.font = { size: 12, bold: true, color: { argb: 'FF111827' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
        cell.border = { top: { style: 'thin', color: { argb: 'FF9CA3AF' } } };
        row.height = 30;
      }

      else if (r.type === 'table-col-headers') {
        tableStripingIndex = 0;
        row = worksheet.addRow(['', ...(r.headers || [])]);
        row.height = 24;
        row.eachCell((cell, colNum) => {
          if (colNum > 1) {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4B5563' } }; // Dark Header
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FF374151' } },
              left: { style: 'thin', color: { argb: 'FF374151' } },
              right: { style: 'thin', color: { argb: 'FF374151' } },
              bottom: { style: 'medium', color: { argb: 'FF374151' } }
            };
          }
        });
      }

      else if (r.type === 'table-row') {
        row = worksheet.addRow(['', ...(r.values || [])]);
        const isEve = tableStripingIndex % 2 === 0;
        tableStripingIndex++;

        row.height = 22;

        row.eachCell((cell, colNum) => {
          if (colNum > 1) {
            cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            cell.border = {
              left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
              right: { style: 'thin', color: { argb: 'FFE5E7EB' } },
              bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } }
            };

            if (isEve) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
            }

            const val = (r.values || [])[colNum - 2];

            if (typeof val === 'number') {
              cell.font = { color: { argb: 'FF0369A1' } };
            }
            else if (typeof val === 'boolean') {
              cell.value = val ? 'TRUE' : 'FALSE';
              cell.font = { color: { argb: 'FF7C3AED' }, bold: true };
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
          }
        });
      }
      else {
        row = worksheet.addRow([]);
      }
      row.outlineLevel = r.level;
    });

    worksheet.getColumn(1).width = 60;
    worksheet.getColumn(2).width = 40;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "GroundTruth_YAML_Style.xlsx");
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    e.target.value = '';

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        if (!buffer) return;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet("Data") || workbook.getWorksheet(1);
        if (!worksheet) {
          alert("Invalid Excel File: Could not find data sheet");
          return;
        }

        console.log("Importing Sheet:", worksheet.name);

        let newData: any = undefined;
        let stack: { obj: any; level: number; key?: string }[] = [];
        let tableCtx: { headers: string[]; array: any[] } | null = null;
        let lastObj: any = null;

        const getIndentLevel = (row: ExcelJS.Row, text: string) => {
          let spaces = 0;
          for (let char of text) {
            if (char === ' ') spaces++;
            else break;
          }
          if (spaces > 0) return Math.floor(spaces / 2);
          return (row.outlineLevel || 0);
        };

        const initRoot = (isArray: boolean) => {
          if (newData !== undefined) return;
          newData = isArray ? [] : {};
          stack = [{ obj: newData, level: -1 }];
          lastObj = newData;
        };

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;

          const cell1 = row.getCell(1);
          const cell2 = row.getCell(2);

          let label = cell1.text || "";
          let val: any = cell2.value;

          const hasInputLabel = label.trim().length > 0;

          if (tableCtx && hasInputLabel) {
            const looksLikeID = !isNaN(Number(label.trim()));
            console.warn(`Row ${rowNumber}: Found label '${label}' inside table context. Ending table.`);

            if (looksLikeID && (val !== null && val !== undefined)) {
              const confirmFix = window.confirm(
                `Row ${rowNumber} seems to have data in the first column ('${label}').\n` +
                `Inside a table, the first column (Structure) should be empty.\n` +
                `Did you mean to put '${label}' as the value for '${tableCtx.headers[0]}'?\n\n` +
                `Click OK to treat it as a data row (Auto-Fit).\n` +
                `Click Cancel to treat it as a new Section (End Table).`
              );
              if (confirmFix) {
                const rowObj: any = {};
                rowObj[tableCtx.headers[0]] = Number(label.trim());
                if (tableCtx.headers[1]) {
                  if (typeof val === 'object' && 'text' in val) val = val.text;
                  rowObj[tableCtx.headers[1]] = val;
                }
                tableCtx.headers.slice(2).forEach((h, i) => {
                  const c = row.getCell(i + 3);
                  let v = c.value;
                  if (v && typeof v === 'object' && 'text' in v) v = (v as any).text;
                  rowObj[h] = v;
                });

                tableCtx.array.push(rowObj);
                return;
              }
            }
          }

          if (hasInputLabel) {
            tableCtx = null;
            const level = getIndentLevel(row, label);
            const key = label.trim();

            initRoot(false);

            while (stack.length > 1 && stack[stack.length - 1].level >= level) {
              stack.pop();
            }

            const parentCtx = stack[stack.length - 1];
            const parent = parentCtx.obj;

            let effectiveKey: string | number = key;
            if (Array.isArray(parent)) effectiveKey = parent.length;

            const isValueEmpty = val === undefined || val === null || (typeof val === 'string' && val.trim() === '');

            if (!isValueEmpty) {
              if (typeof val === 'object' && 'text' in val) val = val.text;
              if (String(val).toUpperCase() === 'TRUE') val = true;
              if (String(val).toUpperCase() === 'FALSE') val = false;

              if (Array.isArray(parent)) parent.push(val);
              else parent[effectiveKey] = val;
            } else {
              const newObj: any = {};
              if (Array.isArray(parent)) parent.push(newObj);
              else parent[effectiveKey] = newObj;

              stack.push({ obj: newObj, level, key: String(effectiveKey) });
              lastObj = newObj;
            }

          } else {
            if (newData === undefined) initRoot(true);

            if (!tableCtx) {
              const headers: string[] = [];
              row.eachCell((cell, colNum) => {
                if (colNum > 1) headers.push(cell.text.trim());
              });

              if (headers.length > 0) {
                const parentCtx = stack[stack.length - 1];
                let targetArray: any[] = [];

                if (parentCtx && parentCtx.obj === lastObj && !Array.isArray(parentCtx.obj)) {
                  targetArray = [];
                  if (stack.length >= 2) {
                    const gp = stack[stack.length - 2].obj;
                    const key = parentCtx.key;
                    if (key !== undefined) {
                      if (Array.isArray(gp)) gp[Number(key)] = targetArray;
                      else gp[key] = targetArray;
                    }
                  } else {
                    newData = targetArray;
                  }
                  parentCtx.obj = targetArray;
                } else if (Array.isArray(parentCtx.obj)) {
                  targetArray = parentCtx.obj;
                }

                tableCtx = { headers, array: targetArray };
              }
            } else {
              const rowObj: any = {};
              let hasData = false;

              tableCtx.headers.forEach((h, idx) => {
                const cell = row.getCell(idx + 2);
                let v = cell.value;
                if (v && typeof v === 'object' && 'text' in v) v = (v as any).text;

                if (v !== undefined && v !== null && String(v).trim() !== '') {
                  if (String(v).toUpperCase() === 'TRUE') v = true;
                  if (String(v).toUpperCase() === 'FALSE') v = false;
                  rowObj[h] = v;
                  hasData = true;
                }
              });

              if (hasData) {
                tableCtx.array.push(rowObj);
              }
            }
          }
        });

        if (newData) {
          setParsedData(newData);
          setJsonInput(JSON.stringify(newData, null, 2));
          setYamlInput(yaml.dump(newData));
          setError(null);
          alert("Import Successful!");
        } else {
          alert("No data found in file.");
        }

      } catch (err) {
        console.error("Import Error", err);
        alert("Error parsing file");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="container">
      <header className="app-header">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="app-title">Ground Truth Generator</h1>
            <p className="app-subtitle">Instantly generate a UI from any JSON structure</p>
          </div>
          <button onClick={handleExportExcel} className="btn btn-primary" title="Download Excel">
            <Download size={18} /> Export Excel
          </button>
          <label className="btn btn-secondary ml-2 cursor-pointer" title="Import Excel">
            <Upload size={18} className="mr-1" /> Import Excel
            <input
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={handleImportExcel}
            />
          </label>
        </div>
      </header>


      {/* Tab Navigation */}
      <div className="tab-nav">
        <button
          className={`tab-btn ${activeTab === 'json' ? 'active' : ''}`}
          onClick={() => setActiveTab('json')}
        >
          <Braces size={16} /> JSON
        </button>
        <button
          className={`tab-btn ${activeTab === 'yaml' ? 'active' : ''}`}
          onClick={() => {
            if (activeTab !== 'yaml') {
              try {
                const obj = JSON.parse(jsonInput);
                setYamlInput(yaml.dump(obj));
              } catch (e) {
                setYamlInput("# Invalid JSON - cannot convert");
              }
            }
            setActiveTab('yaml');
          }}
        >
          <AlignLeft size={16} /> YAML (Excel Alternative)
        </button>
        <button
          className={`tab-btn ${activeTab === 'form' ? 'active' : ''}`}
          onClick={() => setActiveTab('form')}
        >
          <Layout size={16} /> Generated Form
        </button>
      </div>

      <div className={`main-layout ${activeTab ? 'tabbed-mode' : ''}`}>
        {/* Left Panel: JSON Input */}
        {(activeTab === 'json') && (
          <div className="card editor-section" style={{ width: '100%', height: '100%' }}>
            <div className="card-header">
              <div className="card-title">
                <Braces size={18} /> JSON Input
              </div>
              <div className="card-actions">
                <button
                  onClick={() => setWordWrap(!wordWrap)}
                  className={`btn ${wordWrap ? 'btn-active' : ''}`}
                  title="Toggle Word Wrap"
                >
                  <WrapText size={16} />
                </button>
                <button onClick={formatJson} className="btn" title="Format JSON (Prettify)">
                  <AlignLeft size={16} /> Format
                </button>
                <button onClick={() => setJsonInput('{}')} className="btn btn-danger-icon" title="Clear All">
                  <RotateCcw size={16} />
                </button>
              </div>
            </div>

            <div className="card-content">
              <textarea
                value={jsonInput}
                onChange={handleJsonInputChange}
                onBlur={formatJson}
                className={`json-input ${wordWrap ? 'wrap' : ''}`}
                spellCheck={false}
                placeholder="Paste your JSON here..."
              />
              {error && (
                <div className="error-toast">
                  <AlertCircle size={18} />
                  <span>{error}</span>
                </div>
              )}
            </div>
          </div>
        )}

        {/* YAML Tab */}
        {(activeTab === 'yaml') && (
          <div className="card editor-section" style={{ width: '100%', height: '100%' }}>
            <div className="card-header" style={{ backgroundColor: '#f0fdf4' }}>
              <div className="card-title text-green-800">
                <AlignLeft size={18} /> YAML Editor
              </div>
              <div className="card-actions">
                <button onClick={() => {
                  try {
                    const obj = yaml.load(yamlInput);
                    setJsonInput(JSON.stringify(obj, null, 2));
                    setParsedData(obj as any);
                    alert("Synced to Form!");
                  } catch (e) {
                    alert("Invalid YAML");
                  }
                }} className="btn btn-primary-glass text-green-700 border-green-200 hover:bg-green-100">
                  <Check size={16} /> Apply
                </button>
              </div>
            </div>
            <div className="card-content">
              <textarea
                value={yamlInput}
                onChange={(e) => {
                  setYamlInput(e.target.value);
                  try {
                    const obj = yaml.load(e.target.value);
                    if (obj) {
                      setParsedData(obj as any);
                      setJsonInput(JSON.stringify(obj, null, 2));
                      setError(null);
                    }
                  } catch (err) {
                    // Don't error on every keystroke
                  }
                }}
                className="json-input"
                style={{ fontFamily: 'monospace', color: '#166534' }}
                spellCheck={false}
                placeholder="Paste YAML here..."
              />
            </div>
          </div>
        )}

        {/* Right Panel: Form Preview */}
        {(activeTab === 'form') && (
          <div className="card form-section" style={{ width: '100%', height: '100%' }}>
            <div className="card-header">
              <div className="card-title">
                <Check size={18} /> Generated Form
              </div>

              <div className="flex gap-2">
                <button
                  onClick={() => setForceExpandState(true)}
                  className="btn btn-sm btn-ghost hover:bg-slate-700 text-xs px-2"
                  title="Expand All"
                >
                  <Plus size={14} className="mr-1" /> Expand All
                </button>
                <button
                  onClick={() => setForceExpandState(false)}
                  className="btn btn-sm btn-ghost hover:bg-slate-700 text-xs px-2"
                  title="Collapse All"
                >
                  <Minus size={14} className="mr-1" /> Collapse All
                </button>
              </div>

              {error && (
                <div className="error-badge">
                  <AlertCircle size={14} />
                  <span>Invalid JSON</span>
                </div>
              )}
            </div>
            <div className="card-content form-content">
              {parsedData === null ? (
                <div className="empty-state">
                  <AlertCircle size={48} style={{ marginBottom: '1rem', opacity: 0.5 }} />
                  <h3>No Data</h3>
                  <p>Enter valid JSON to see the form.</p>
                </div>
              ) : (
                <div>
                  {typeof parsedData === 'object' && !Array.isArray(parsedData) ? (
                    Object.entries(parsedData).map(([key, val]) => (
                      <DynamicField
                        key={key}
                        label={key}
                        value={val}
                        path={[key]}
                        onChange={handleFormChange}
                        onRemove={handleFormRemove}
                        forceExpandState={forceExpandState}
                        highlightedPath={highlightedPath}
                        setHighlightedPath={setHighlightedPath}
                      />
                    ))
                  ) : Array.isArray(parsedData) ? (
                    <DynamicField
                      label="Root Array"
                      value={parsedData}
                      path={[]}
                      onChange={handleFormChange}
                      onRemove={handleFormRemove}
                      forceExpandState={forceExpandState}
                      highlightedPath={highlightedPath}
                      setHighlightedPath={setHighlightedPath}
                    />
                  ) : (
                    <DynamicField
                      label="Root Value"
                      value={parsedData}
                      path={[]}
                      onChange={handleFormChange}
                      forceExpandState={forceExpandState}
                      highlightedPath={highlightedPath}
                      setHighlightedPath={setHighlightedPath}
                    />
                  )}

                  {typeof parsedData === 'object' && !Array.isArray(parsedData) && (
                    <RootObjectAdder
                      onChange={handleFormChange}
                      rootData={parsedData}
                      setHighlightedPath={setHighlightedPath}
                    />
                  )}

                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
