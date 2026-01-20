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
  Download
} from 'lucide-react';
import * as XLSX from 'xlsx';

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
  onRemove
}: {
  label: string | number,
  value: Value,
  path: (string | number)[],
  onChange: (path: (string | number)[], val: any) => void,
  onRemove?: (path: (string | number)[]) => void
}) => {

  const isArray = Array.isArray(value);
  const isObject = value !== null && typeof value === 'object' && !isArray;
  const isPrimitive = !isArray && !isObject;

  // State for adding new field to Object
  const [isAdding, setIsAdding] = useState(false);
  const [newKey, setNewKey] = useState("");
  const [newType, setNewType] = useState("string");

  const handleAddArrayItem = () => {
    if (!isArray) return;
    let newItem: any = "";
    if (value.length > 0) {
      // Clone structure
      const cloneStructure = (item: any): any => {
        if (Array.isArray(item)) return [];
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
    setIsAdding(false);
    setNewKey("");
    setNewType("string");
  };

  const getTypeLabel = () => {
    if (isArray) return 'Array';
    if (isObject) return 'Object';
    return typeof value;
  };

  return (
    <div className={`field-container ${isPrimitive ? 'input-group' : ''}`}>
      <div className="field-header">
        <label className="field-label">
          <span className="font-medium text-slate-200">{label}</span>
          <span className="type-badge">{getTypeLabel()}</span>
        </label>

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
              className="input"
            />
          ) : (
            <input
              type="text"
              value={value as string}
              onChange={(e) => onChange(path, e.target.value)}
              className="input"
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
              onRemove={onRemove ? undefined : (p) => {
                // Allow removing keys from top-level or nested objects
                // Logic: parent component passed us 'onChange', so we can't easily call 'onRemove' of parent.
                // Actually, we need to pass a callback to children to remove themselves.
                // The recursive DynamicField call needs onRemove logic.
                // We don't have direct access to 'modify parent', but we have 'onChange'.
                // So we can implement 'onRemove' for children by using 'removeValue' logic at root?
                // No, simply:
                const newObj = { ...value };
                delete newObj[key];
                onChange(path, newObj);
              }}
            />
          ))}

          {/* Add Field UI */}
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
        <div className="nested-object">
          {value.map((item, index) => (
            <div key={index} className="array-item">
              <DynamicField
                label={`Item ${index + 1}`}
                value={item}
                path={[...path, index]}
                onChange={onChange}
                onRemove={(p) => onRemove ? onRemove(p) : onChange(path, value.filter((_, i) => i !== index))}
              />
            </div>
          ))}
          <button onClick={handleAddArrayItem} className="btn btn-primary" style={{ width: '100%', marginTop: '0.5rem', justifyContent: 'center' }}>
            <Plus size={16} /> Add Item
          </button>
        </div>
      )}
    </div>
  );
};

// --- Main App ---

function App() {
  const [jsonInput, setJsonInput] = useState<string>('{\n  "title": "My Awesome Form",\n  "isActive": true,\n  "count": 42,\n  "tags": ["hero", "dark-mode", "react"],\n  "author": {\n    "name": "Jane Doe",\n    "email": "jane@example.com"\n  },\n  "features": [\n    { "id": 1, "name": "Login" },\n    { "id": 2, "name": "Dashboard" }\n  ]\n}');
  const [parsedData, setParsedData] = useState<Value | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [wordWrap, setWordWrap] = useState<boolean>(false);

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
  const generateHierarchy = (data: any) => {
    const rows: { row: any, level: number }[] = [];

    // Recursive traversal to build rows with depth info
    const traverse = (key: string, value: any, level: number) => {
      const indent = "    ".repeat(level);
      const displayKey = indent + key;
      const type = Array.isArray(value) ? 'Array' : (value === null ? 'null' : typeof value);

      const rowItem = {
        "Structure": displayKey,
        "Value": (value === null || typeof value !== 'object') ? value : "",
        "Type": type
      };

      // Push current row
      rows.push({ row: rowItem, level: level });

      // Recurse if object
      if (value !== null && typeof value === 'object') {
        Object.entries(value).forEach(([k, v]) => {
          traverse(k, v, level + 1);
        });
      }
    };

    if (typeof data === 'object' && data !== null) {
      if (Array.isArray(data)) {
        traverse('Root Array', data, 0);
      } else {
        // Iterate root keys directly to avoid extra indent for root object
        Object.entries(data).forEach(([k, v]) => traverse(k, v, 0));
      }
    } else {
      traverse('Root', data, 0);
    }

    return rows;
  };

  const handleExportExcel = () => {
    if (!parsedData) return;

    // 1. Generate Data with Levels
    const complexRows = generateHierarchy(parsedData);

    // 2. Extract plain data for the sheet
    const sheetData = complexRows.map(r => r.row);
    const worksheet = XLSX.utils.json_to_sheet(sheetData);

    // 3. Apply Grouping (Outlining)
    // We must initialize the !rows array. 
    // It filters sparse arrays, so we should fill it efficiently.
    const rowProps: XLSX.RowInfo[] = [];

    // Header Row (Level 0) - usually level 0 means visible/top-level
    rowProps.push({ level: 0 });

    // Data Rows
    complexRows.forEach((item) => {
      rowProps.push({ level: item.level });
    });

    worksheet['!rows'] = rowProps;

    // 4. Configure Outline Direction
    // summaryBelow: false means parents are ABOVE children (standard Tree view).
    worksheet['!outline'] = { summaryBelow: false };

    // 5. Formatting (Widths)
    const maxKeyLen = sheetData.reduce((max, r) => Math.max(max, (r.Structure || "").length), 15);
    const maxValLen = sheetData.reduce((max, r) => Math.max(max, String(r.Value || "").length), 15);
    worksheet['!cols'] = [
      { wch: maxKeyLen + 5 },
      { wch: maxValLen + 5 },
      { wch: 10 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Hierarchy");

    // Use 'xlsx' book type explicitly to preserve features
    XLSX.writeFile(workbook, "hierarchical_grouped.xlsx", { bookType: 'xlsx', type: 'binary' });
  };

  return (
    <div className="container">
      <header className="app-header">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="app-title">Dynamic JSON Form</h1>
            <p className="app-subtitle">Instantly generate a UI from any JSON structure</p>
          </div>
          <button onClick={handleExportExcel} className="btn btn-primary" title="Download Excel">
            <Download size={18} /> Export Excel
          </button>
        </div>
      </header>

      <div className="main-layout">
        {/* Left Panel: JSON Input */}
        <div className="card editor-section">
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

        {/* Right Panel: Generated Form */}
        <div className="card preview-section">
          <div className="card-header">
            <div className="card-title text-white">
              <Check size={18} /> Generated Form
            </div>
            <div className="text-xs text-slate-400">Live Preview</div>
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
                    />
                  ))
                ) : Array.isArray(parsedData) ? (
                  <DynamicField
                    label="Root Array"
                    value={parsedData}
                    path={[]}
                    onChange={handleFormChange}
                    onRemove={handleFormRemove}
                  />
                ) : (
                  <DynamicField
                    label="Root Value"
                    value={parsedData}
                    path={[]}
                    onChange={handleFormChange}
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
                  <RootObjectAdder onChange={handleFormChange} rootData={parsedData} />
                )}

              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// --- Helper Component for Root Object Adder ---
const RootObjectAdder = ({ onChange, rootData }: { onChange: any, rootData: any }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [newKey, setNewKey] = useState("");
  const [newType, setNewType] = useState("string");

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
    // Root path is empty array. Adding key means creating new object with that key..
    // Actually we just passed 'onChange' which expects 'path'.
    // To add a key 'foo' to root, path is ['foo'].
    onChange([newKey], initialValue);
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
}

export default App;
