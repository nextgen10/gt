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
    // @ts-ignore
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
    // @ts-ignore
    newArr[head] = removeValue(obj[head], tail);
    return newArr;
  }
  return { ...obj, [head]: removeValue(obj[head], tail) };
};

// --- Helper: Flatten JSON for Excel ---


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
  forceExpandState, // null = no force, true = expand all, false = collapse all
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

  // Determine if we should render as a table (Array of Objects)
  const isTableMode = isArray && value.length > 0 &&
    value.every((item: any) => item !== null && typeof item === 'object' && !Array.isArray(item));

  // State for adding new field to Object
  const [isAdding, setIsAdding] = useState(false);
  const [newKey, setNewKey] = useState("");
  const [newType, setNewType] = useState("string");

  // State for collapse/expand
  // Default: Root (level 0) expanded, others collapsed
  const [isExpanded, setIsExpanded] = useState(level === 0);

  // React to Force Expand/Collapse signals
  useEffect(() => {
    if (forceExpandState !== null && forceExpandState !== undefined) {
      setIsExpanded(forceExpandState);
    }
  }, [forceExpandState]);

  // State for editing key name
  const [isEditingKey, setIsEditingKey] = useState(false);
  const [editedKey, setEditedKey] = useState(String(label));

  useEffect(() => {
    setEditedKey(String(label));
  }, [label]);


  const handleAddArrayItem = () => {
    if (!isArray) return;
    let newItem: any = "";
    if (value.length > 0) {
      // Clone structure
      const cloneStructure = (item: any): any => {
        if (Array.isArray(item)) {
          // If the array has items, we want to preserve the structure of the inner items
          // by creating an array with ONE blank template item.
          // This ensures nested arrays (like 'countries' inside 'regions') don't become empty []
          // which would lose the schema effectively blocking further additions.
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
      // Also expand the parent (current array) so the new item is visible? 
      // Current component is the array field itself. It must be expanded to see children.
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
              // Allow toggling by clicking the label area for non-primitives
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

            {/* Editable Label */}
            {isEditingKey ? (
              <input
                autoFocus
                type="text"
                value={editedKey}
                onChange={(e) => setEditedKey(e.target.value)}
                onBlur={() => {
                  // Trigger rename
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
                onClick={(e) => e.stopPropagation()} // Prevent collapse
              />
            ) : (
              <span
                className="font-medium text-slate-200 hover:text-white cursor-text border-b border-transparent hover:border-slate-500 transition-colors"
                title="Click to rename key"
                onClick={(e) => {
                  // Only allow editing keys if onRename is present (i.e. we are in an object)
                  if (onRename) {
                    e.preventDefault();
                    setIsEditingKey(true);
                  }
                }}
              >
                {label}
              </span>
            )}

            {/* Only show badge if header is shown */}
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

      {/* Mini-Header for toggle when standard header is hidden */}
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

                  // NEW: Pass onRename prop
                  onRename={(newKeyName) => {
                    if (newKeyName === key) return;
                    if (!newKeyName.trim()) return;
                    // Determine if key already exists?
                    if (Object.keys(value).includes(newKeyName)) {
                      alert("Key already exists!");
                      return;
                    }

                    const newObj = { ...value };
                    // Preserve order if possible, or just add new key/value
                    // To rename, we delete old and add new with same value.
                    // Ideally we want to keep position.
                    const keys = Object.keys(newObj);

                    // Reconstruct object to maintain order
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

              {/* ... existing Add Field UI ... */}
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
                      {/* Gather all unique keys from all objects to form headers */}
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


// --- Helper Component for Root Object Adder ---
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
      // Logic: If initial load, maybe default expandable? 
      // But user requested "collapse everything in form when we add new json".
      // So on initial load or subsequent updates, we might want to default to collapse?
      // Actually usually "add new json" implies manual paste.
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
      // Auto-collapse on new valid JSON
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

  // --- Helper: Generate Hierarchical Rows for Excel ---


  const handleExportExcel = async () => {
    if (!parsedData) return;

    // --- INTERNAL GENERATOR ---
    type ExcelRow = {
      type: 'field' | 'table-header' | 'table-col-headers' | 'table-row';
      path: string;
      label: string; // Display Key
      value?: any;
      headers?: string[];
      values?: any[];
      level: number;
    };

    const rows: ExcelRow[] = [];

    // Traverse: obj, fullPath, displayKey, level
    const traverse = (obj: any, path: string, key: string, level: number) => {
      // 1. Check for Table (Array of Objects)
      const isTable = Array.isArray(obj) && obj.length > 0 &&
        obj.every((i: any) => i && typeof i === 'object' && !Array.isArray(i));

      if (isTable) {
        // A. Table Block Header (Key Name)
        rows.push({ type: 'table-header', path: path, label: key, level });

        // B. Calculate Columns
        const keys = new Set<string>();
        obj.forEach((row: any) => Object.keys(row).forEach(k => keys.add(k)));
        const colHeaders = Array.from(keys);

        // C. Column Headers Row (Indent + 1)
        // We don't really have a label for this row, it uses col headers.
        rows.push({ type: 'table-col-headers', path: '', label: '', headers: colHeaders, level: level + 1 });

        // D. Data Rows
        obj.forEach((rowObj: any, rowIndex: number) => {
          const rowVals = colHeaders.map(k => {
            const val = rowObj[k];
            if (val && typeof val === 'object') {
              return Array.isArray(val) ? `[Array(${val.length})]` : '[Object]';
            }
            return val;
          });
          // Table Rows are at level + 1 (indented under the table header)
          rows.push({ type: 'table-row', path: '', label: '', values: rowVals, level: level + 1 });

          // E. Recurse for Complex Children
          Object.entries(rowObj).forEach(([k, v]) => {
            if (v && typeof v === 'object') {
              const childPath = path ? `${path}.${rowIndex}.${k}` : `${rowIndex}.${k}`;
              // Recursive tables/objects start at level + 2 (under the row, conceptually)
              // Or maybe just level + 1 if we treat them as siblings to the row content?
              // "YAML" style:
              // - Item 1
              //   key: val
              //   regions:
              //     ...
              // Let's use level + 1 relative to the Row, so Level + 2 total.
              // But we need a parent label for the deep object? Currently we use 'k'.
              // Wait, `traverse` expects `key`.
              // We need to render the key `k` as the header for the next block.
              traverse(v, childPath, k, level + 2);
            }
          });
        });
        return;
      }

      // 2. Standard Traversal
      if (Array.isArray(obj)) {
        obj.forEach((v: any, i: number) => {
          // For Arrays, key is usually index, but in YAML lists are hyphenated "-".
          // We'll use "- (Index)" or just "-".
          traverse(v, path ? `${path}.${i}` : `${i}`, `Item ${i + 1}`, level);
        });
        return;
      }

      if (obj && typeof obj === 'object') {
        // If this is a nested object, we need a separate "Header Row" for the object Key
        // UNLESS it's the root or we are already inside a recursive call that pushed the key?
        // In `traverse`, we are processing `obj`. `key` is passed in.
        // If it's the Root, we might not want a row.
        // If it's a nested object field, we usually want: `Key:` row, then children indented.

        if (key && key !== 'Root') {
          rows.push({ type: 'table-header', path: path, label: key, level });
          level++; // Indent children
        }

        Object.entries(obj).forEach(([k, v]) => {
          traverse(v, path ? `${path}.${k}` : k, k, level);
        });
        return;
      }

      // 3. Primitive Field
      rows.push({ type: 'field', path: path, label: key, value: obj, level });
    };

    // Start
    traverse(parsedData, '', '', 0);

    // --- RENDER TO EXCEL ---
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data", { views: [{ showGridLines: false }] });

    // Header Row (Main)
    const mainHeader = worksheet.addRow(['Structure', 'Value']);
    mainHeader.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    mainHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
    mainHeader.height = 30;

    // Configure Outline
    worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };

    // Process Rows
    rows.forEach(r => {
      let row: ExcelJS.Row;

      // Indentation String (2 spaces per level)
      const indentStr = "  ".repeat(r.level);

      if (r.type === 'field') {
        row = worksheet.addRow([indentStr + r.label, r.value]); // Indented Label

        const pathCell = row.getCell(1);
        pathCell.font = { color: { argb: 'FF374151' }, bold: true };
        pathCell.alignment = { vertical: 'middle' }; // No alignment indent, using spaces

        const valCell = row.getCell(2);
        valCell.alignment = { vertical: 'middle', wrapText: true };

        if (typeof r.value === 'boolean') {
          valCell.value = r.value ? 'TRUE' : 'FALSE';
          valCell.font = { color: { argb: 'FF7C3AED' }, bold: true };
          valCell.dataValidation = { type: 'list', allowBlank: false, formulae: ['"TRUE,FALSE"'] };
        } else if (typeof r.value === 'number') {
          valCell.font = { color: { argb: 'FF0284C7' } };
          valCell.alignment = { horizontal: 'left' };
        }

        row.eachCell(c => c.border = { bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } } });
      }

      else if (r.type === 'table-header') {
        // Just the Key Name (Section Header)
        row = worksheet.addRow([indentStr + r.label]);
        const cell = row.getCell(1);
        cell.font = { size: 12, bold: true, color: { argb: 'FF111827' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
        row.height = 25;
      }

      else if (r.type === 'table-col-headers') {
        // [Empty/Indent, ...Headers]
        // Alignment: Col 1 is empty indentation. Headers start at Col 2.
        // Actually, if we use space-indentation, we can just put empty string in Col 1.
        row = worksheet.addRow(['', ...(r.headers || [])]);
        row.eachCell((cell, colNum) => {
          if (colNum > 1) {
            cell.font = { bold: true, color: { argb: 'FF4B5563' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
            cell.border = { bottom: { style: 'medium', color: { argb: 'FFD1D5DB' } } };
          }
        });

        // Ensure indentation of the row itself matches hierarchy? 
        // Excel outline handles the collapse. Visually, the headers are distinct.
      }

      else if (r.type === 'table-row') {
        // [Empty/Indent, ...Values]
        row = worksheet.addRow(['', ...(r.values || [])]);
        row.eachCell((cell, colNum) => {
          if (colNum > 1) {
            const val = (r.values || [])[colNum - 2];
            cell.value = val;
            cell.border = { bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } } };

            if (typeof val === 'number') cell.font = { color: { argb: 'FF0284C7' } };
            if (typeof val === 'boolean') {
              cell.value = val ? 'TRUE' : 'FALSE';
              cell.font = { color: { argb: 'FF7C3AED' }, bold: true };
              cell.dataValidation = { type: 'list', allowBlank: false, formulae: ['"TRUE,FALSE"'] };
            }
          }
        });
      }
      else {
        row = worksheet.addRow([]);
      }

      // --- APPLY GROUPING ---
      row.outlineLevel = r.level;
    });

    // Auto widths
    worksheet.getColumn(1).width = 60;
    worksheet.getColumn(2).width = 40;

    // Write
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "GroundTruth_YAML_Style.xlsx");
  };



  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const buffer = evt.target?.result as ArrayBuffer;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.getWorksheet("Hierarchy");
      if (!worksheet) {
        alert("Invalid Excel File: Missing 'Hierarchy' sheet");
        return;
      }

      // Reconstruct Data
      const setDeepValue = (obj: any, pathStr: string, val: any) => {
        if (!pathStr) return obj;
        const path = pathStr.split('.').map(p => isNaN(Number(p)) ? p : Number(p));
        // Inline robust setValue wrapper
        const setRobust = (o: any, p: (string | number)[], v: any): any => {
          if (p.length === 0) return v;
          const [h, ...t] = p;
          const current = o || {};
          const isArr = Array.isArray(current);
          const nextIsArr = t.length > 0 && typeof t[0] === 'number';
          if (isArr) {
            const newArr = [...current];
            const idx = Number(h);
            if (newArr[idx] === undefined) newArr[idx] = nextIsArr ? [] : {};
            newArr[idx] = setRobust(newArr[idx], t, v);
            return newArr;
          }
          const newObj = { ...current };
          if (newObj[h] === undefined) newObj[h] = nextIsArr ? [] : {};
          newObj[h] = setRobust(newObj[h], t, v);
          return newObj;
        };
        return setRobust(obj, path, val);
      };

      // 0. Parse Headers explicitly
      const colMap: Record<string, number> = {};
      const dataCols: number[] = [];

      const headerRow = worksheet.getRow(1);
      headerRow.eachCell((cell, colNum) => {
        const headerVal = cell.value?.toString().trim() || "";
        if (headerVal) {
          colMap[headerVal] = colNum;
          if (!['Path', 'Structure', 'Value', 'Type'].includes(headerVal)) {
            dataCols.push(colNum);
          }
        }
      });

      const getVal = (row: ExcelJS.Row, headerName: string): any => {
        const idx = colMap[headerName];
        if (!idx) return undefined;
        return row.getCell(idx).value;
      };

      let newData: any = {};

      // Determine Root (Array vs Object)
      let isRootArray = false;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const pathVal = getVal(row, 'Path')?.toString();
        // If first path starts with "0.", it's an array
        if (pathVal && !isNaN(Number(pathVal.split('.')[0]))) {
          isRootArray = true;
        }
      });
      newData = isRootArray ? [] : {};

      // 1. SEQUENTIAL PROCESSING
      // We process the file Top-to-Bottom.
      // We maintain "Current Context" (Are we inside a table? Which index are we on?)
      // This completely ignores the explicit indices in the 'Path' column for Table Rows,
      // trusting the visual file order instead.

      let currentTablePath: string | null = null;
      let currentTableIndex: number = 0;

      // Indentation Stack for Standard Fields: [{ path: "", indent: -1 }]
      // Used to infer parent of new rows based on "Structure" indentation
      let indentStack: { path: string, indent: number }[] = [{ path: "", indent: -1 }];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        let path = getVal(row, 'Path')?.toString() || "";
        let type = getVal(row, 'Type')?.toString();

        // A. DETECT TABLE CONTEXT SWITCH
        if (type === 'Table') {
          // Starting a new Table Block
          currentTablePath = path;
          currentTableIndex = 0; // Reset counter
          return; // This row is just a header/container, no value
        }

        if (type === 'TableHeader') return; // Skip

        // B. DETECT TABLE ROW (Explicit or Inferred)
        // If Type is 'TableRow', it is part of the current table.
        // If Path is empty but we are in a table context, it is a new inserted row.
        // If Path exists but matches the table prefix, it is an edited row (we ignore its old index).

        let isTableRow = false;

        if (type === 'TableRow') {
          isTableRow = true;
        } else if (!type && !path && currentTablePath) {
          // Empty row inside a table block -> New Item
          isTableRow = true;
        } else if (currentTablePath && path.startsWith(currentTablePath + '.')) {
          // It has a path, check if it looks like a table item
          // Verify it is immediate child logic? 
          // Actually, standard "Array of Objects" export creates paths like "root.items.0"
          // If currentTablePath is "root.items", this checks out.
          isTableRow = true;
        }

        // C. PROCESS ROW
        if (isTableRow && currentTablePath) {
          // --- TABLE MODE ---
          // We FORCE the path to use our sequential index.
          const effectivePath = `${currentTablePath}.${currentTableIndex}`;

          // Read all dynamic columns
          dataCols.forEach(colIdx => {
            const key = Object.keys(colMap).find(k => colMap[k] === colIdx);
            if (!key) return;

            const cellVal = row.getCell(colIdx).value;
            let finalVal: any = cellVal;

            if (typeof cellVal === 'object' && cellVal !== null && 'text' in cellVal) {
              finalVal = (cellVal as any).text;
            }

            if (finalVal !== undefined && finalVal !== null && finalVal !== '') {
              if (String(finalVal).toLowerCase() === 'true') finalVal = true;
              else if (String(finalVal).toLowerCase() === 'false') finalVal = false;
              else if (!isNaN(Number(finalVal)) && String(finalVal).trim() !== '') finalVal = Number(finalVal);

              newData = setDeepValue(newData, `${effectivePath}.${key}`, finalVal);
            }
          });

          // Increment for NEXT row
          currentTableIndex++;

        } else {
          // --- STANDARD MODE (Primitive / Individual Field) ---

          // 1. Calculate Indentation logic
          const structureRaw = getVal(row, 'Structure')?.toString() || "";

          // Determine indent based on leading hyphens (---- per level)
          // match start of string with one or more hyphens or spaces (backward compat?)
          // Let's stick to user request: 4 hyphens.
          const leadingMarkersMatch = structureRaw.match(/^[\s-]*/);
          const leadingMarkersCount = leadingMarkersMatch ? leadingMarkersMatch[0].length : 0;

          // If strictly hyphens, we replace them.
          const trimmedKey = structureRaw.replace(/^[\s-]+/, '').trim();

          // 2. Sync Hierarchy Stack
          // If we have an explicit path, we align the stack to it.
          // If not, we align the stack to the indentation.

          if (path) {
            // Existing Row: Trust its path, push to stack
            // Clear stack of deeper/equal items if this is a sibling/uncle
            while (indentStack.length > 1 && indentStack[indentStack.length - 1].indent >= leadingMarkersCount) {
              indentStack.pop();
            }
            indentStack.push({ path: path, indent: leadingMarkersCount });

            // We might have exited the table?
            if (currentTablePath && !path.startsWith(currentTablePath)) {
              currentTablePath = null;
            }

          } else {
            // New Row (Missing Path): Infer from Indentation
            // Pop stack until we find a parent (indent < current)
            while (indentStack.length > 1 && indentStack[indentStack.length - 1].indent >= leadingMarkersCount) {
              indentStack.pop();
            }
            const parent = indentStack[indentStack.length - 1];

            // Construct New Path
            if (parent.path) {
              path = `${parent.path}.${trimmedKey}`;
            } else {
              path = trimmedKey; // Root level
            }

            // Push inferred context
            indentStack.push({ path: path, indent: leadingMarkersCount });
          }

          // If still no path, we can't do anything (orphaned row)
          if (!path) return;

          const explicitType = getVal(row, 'Type')?.toString();
          let finalVal: any = getVal(row, 'Value');

          if (typeof finalVal === 'object' && finalVal !== null && 'text' in finalVal) {
            finalVal = (finalVal as any).text;
          }

          if (explicitType === 'boolean') {
            finalVal = (String(finalVal).toLowerCase() === 'true');
          } else if (explicitType === 'number') {
            finalVal = Number(finalVal);
          } else if (explicitType === 'null') {
            finalVal = null;
          } else {
            if (!explicitType && finalVal !== null && finalVal !== undefined) {
              if (String(finalVal).toLowerCase() === 'true') finalVal = true;
              else if (String(finalVal).toLowerCase() === 'false') finalVal = false;
              else if (!isNaN(Number(finalVal)) && String(finalVal).trim() !== '') finalVal = Number(finalVal);
            }
          }

          newData = setDeepValue(newData, path, finalVal);
        }
      });

      setParsedData(newData);
      setJsonInput(JSON.stringify(newData, null, 2));
      e.target.value = '';
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
            // Convert current JSON to YAML when switching to tab
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
                    // Don't error on every keystroke, just don't update
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

                  {/* Special case: If root is Object, allow adding fields to it directly?
                       The renderer above maps entries. We need to wrap root object in a DynamicField?
                       Actually, the recursive structure handles it best if we treat root as just a value.
                       However, the current App.tsx maps Object.entries manually for the root object 
                       to avoid an extra "Root" wrapper which looks ugly.
                       But this means we can't easily add fields to root.
                       Let's Fix: We should just show an "Add Field" button at the bottom of the Root render loop.
                   */}
                  {typeof parsedData === 'object' && !Array.isArray(parsedData) && (
                    /* We need to reuse the UI logic for adding fields. 
                       Ideally we should refactor "Add Field" into a component or just duplicate it here briefly.
                       For simplicity/cleanliness, let's wrap the root in DynamicField, 
                       BUT with a special "transparent" mode?
                       No, simpler: Just render the Logic here. 
                    */
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
