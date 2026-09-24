import React, {
  useState,
  useRef,
  useMemo,
  useCallback,
  useEffect,
} from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Plus,
  X,
  History,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Trash2,
  Check,
  ArrowDownWideNarrow,
  Lock,
  RotateCcw,
  FileSpreadsheet,
  Clock,
  BarChart3,
  Sigma,
  Download,
  Pencil,
  Database,
  Link2,
  Save,
} from "lucide-react";

// ============================================================
// DESIGN TOKENS
// ============================================================

const T = {
  bg: "#F5F6F2",
  surface: "#FFFFFF",
  border: "#DBDFD5",
  borderStrong: "#C3C8BC",
  ink: "#1D211C",
  muted: "#6E7469",
  faint: "#9BA096",
  amber: "#C77D2E",
  amberBg: "#FBEFDF",
  teal: "#2B5D53",
  tealDark: "#1E453D",
  tealBg: "#E7F0EE",
  red: "#B4483A",
  blue: "#4267A8",
  blueBg: "#EDF2FA",
};

const sans = "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif";
const mono = "ui-monospace, 'SFMono-Regular', Menlo, Consolas, monospace";

const TOP_N = 8;

// ============================================================
// ID HELPERS
// ============================================================

let tabCounter = 1;
let workingSetCounter = 1;
let historyCounter = 1;

function makeTabId() {
  return `tab-${tabCounter++}`;
}

function makeWorkingSetId() {
  return `set-${workingSetCounter++}`;
}

function makeHistoryId() {
  return `hist-${historyCounter++}`;
}

// ============================================================
// DATA HELPERS
// ============================================================

function cloneRows(rows) {
  return rows.map((row, index) => ({
    ...row,
    _rid: `${Date.now()}-${Math.random().toString(36).slice(2)}-${index}`,
  }));
}

function getDisplayColumns(rows) {
  if (!rows.length) return [];

  const keys = new Set();

  rows.forEach((row) => {
    Object.keys(row).forEach((key) => {
      if (key !== "_rid" && key !== "_sourceSet") {
        keys.add(key);
      }
    });
  });

  return [...keys];
}

function detectColumns(rows) {
  if (!rows.length) return [];

  const keys = getDisplayColumns(rows);

  return keys.map((name) => {
    const vals = rows
        .map((r) => r[name])
        .filter((v) => v !== undefined && v !== null && v !== "");

    const isNumeric =
        vals.length > 0 &&
        vals.every(
            (v) => typeof v === "number" && !Number.isNaN(v)
        );

    if (isNumeric) {
      return {
        name,
        kind: "numeric",
        min: Math.min(...vals),
        max: Math.max(...vals),
      };
    }

    const options = Array.from(
        new Set(vals.map((v) => String(v)))
    ).sort((a, b) => a.localeCompare(b));

    return {
      name,
      kind: "categorical",
      options,
    };
  });
}

function emptyFilters(columns) {
  const filters = {};

  columns.forEach((column) => {
    if (column.kind === "categorical") {
      filters[column.name] = {
        kind: "categorical",
        values: [],
      };
    } else {
      filters[column.name] = {
        kind: "numeric",
        operator: "gt",
        value: "",
        value2: "",
      };
    }
  });

  return filters;
}

function ruleMatches(row, colName, rule) {
  const value = row[colName];

  if (!rule) return true;

  if (rule.kind === "categorical") {
    if (!rule.values || rule.values.length === 0) {
      return true;
    }

    return rule.values.includes(String(value));
  }

  if (rule.kind === "numeric") {
    if (
        value === undefined ||
        value === null ||
        value === "" ||
        typeof value !== "number"
    ) {
      return false;
    }

    if (rule.operator === "gt") {
      return rule.value !== "" && value > Number(rule.value);
    }

    if (rule.operator === "lt") {
      return rule.value !== "" && value < Number(rule.value);
    }

    if (rule.operator === "eq") {
      return rule.value !== "" && value === Number(rule.value);
    }

    if (rule.operator === "between") {
      if (rule.value === "" || rule.value2 === "") {
        return true;
      }

      return (
          value >= Number(rule.value) &&
          value <= Number(rule.value2)
      );
    }
  }

  return true;
}

function isRuleActive(rule) {
  if (!rule) return false;

  if (rule.kind === "categorical") {
    return rule.values && rule.values.length > 0;
  }

  if (rule.kind === "numeric") {
    if (rule.operator === "between") {
      return rule.value !== "" && rule.value2 !== "";
    }

    return rule.value !== "";
  }

  return false;
}

function applyPriority(data, filters) {
  const activeCols = Object.keys(filters).filter((column) =>
      isRuleActive(filters[column])
  );

  if (activeCols.length === 0) {
    return {
      result: [...data],
      matchedIds: new Set(),
      activeCols,
    };
  }

  const matched = [];
  const rest = [];

  data.forEach((row) => {
    const isMatch = activeCols.every((column) =>
        ruleMatches(row, column, filters[column])
    );

    if (isMatch) {
      matched.push(row);
    } else {
      rest.push(row);
    }
  });

  return {
    result: [...matched, ...rest],
    matchedIds: new Set(matched.map((row) => row._rid)),
    activeCols,
  };
}

function sumColumn(rows, colName) {
  return rows.reduce(
      (sum, row) =>
          typeof row[colName] === "number"
              ? sum + row[colName]
              : sum,
      0
  );
}

function fmtNum(n) {
  return n.toLocaleString(undefined, {
    maximumFractionDigits: 2,
  });
}

// ============================================================
// COMBINE / JOIN HELPERS
// ============================================================
//
// This implementation treats "join working sets together" as
// combining/stacking their rows into one working dataset.
//
// Example:
//
// Set A: Name, Age
// Set B: Name, Age
//
// Result:
// Name, Age
// ...
//
// If one set has columns the other doesn't have, the missing
// cells are simply blank.
//
// The source sets remain completely untouched.
// ============================================================

function combineWorkingSets(selectedSets) {
  if (!selectedSets.length) return [];

  const allColumns = Array.from(
      new Set(
          selectedSets.flatMap((set) =>
              getDisplayColumns(set.rows)
          )
      )
  );

  const combined = [];

  selectedSets.forEach((set) => {
    set.rows.forEach((row) => {
      const newRow = {
        _rid: `${Date.now()}-${Math.random()
            .toString(36)
            .slice(2)}-${combined.length}`,
        _sourceSet: set.name,
      };

      allColumns.forEach((column) => {
        newRow[column] =
            row[column] === undefined ? "" : row[column];
      });

      combined.push(newRow);
    });
  });

  return combined;
}

// ============================================================
// BUTTON
// ============================================================

function Btn({
               children,
               onClick,
               variant = "ghost",
               disabled,
               icon: Icon,
               style,
             }) {
  const styles = {
    ghost: {
      background: "transparent",
      color: T.ink,
      border: `1px solid ${T.border}`,
    },

    primary: {
      background: T.teal,
      color: "#fff",
      border: `1px solid ${T.teal}`,
    },

    danger: {
      background: "transparent",
      color: T.red,
      border: `1px solid ${T.red}`,
    },

    subtle: {
      background: T.surface,
      color: T.muted,
      border: `1px solid ${T.border}`,
    },
  };

  return (
      <button
          onClick={onClick}
          disabled={disabled}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            padding: "7px 12px",
            borderRadius: 6,
            fontSize: 13,
            fontFamily: sans,
            cursor: disabled ? "default" : "pointer",
            opacity: disabled ? 0.45 : 1,
            transition: "background 120ms",
            whiteSpace: "nowrap",
            ...styles[variant],
            ...style,
          }}
      >
        {Icon && <Icon size={14} />}
        {children}
      </button>
  );
}

// ============================================================
// DROPDOWN
// ============================================================

function SelectBox({
                     value,
                     onChange,
                     options,
                     placeholder,
                     disabled = false,
                   }) {
  return (
      <select
          value={value}
          disabled={disabled}
          onChange={(e) => onChange(e.target.value)}
          style={{
            width: "100%",
            padding: "8px 9px",
            borderRadius: 6,
            border: `1px solid ${T.border}`,
            background: T.surface,
            color: value ? T.ink : T.faint,
            fontFamily: sans,
            fontSize: 13,
            outline: "none",
          }}
      >
        {placeholder && (
            <option value="">
              {placeholder}
            </option>
        )}

        {options.map((option) => (
            <option
                key={option.value}
                value={option.value}
                disabled={option.disabled}
            >
              {option.label}
            </option>
        ))}
      </select>
  );
}

// ============================================================
// MULTI SELECT
// ============================================================

function MultiSelect({
                       options,
                       selected,
                       onChange,
                       disabled,
                       placeholder,
                     }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    function onDoc(e) {
      if (
          ref.current &&
          !ref.current.contains(e.target)
      ) {
        setOpen(false);
      }
    }

    document.addEventListener("mousedown", onDoc);

    return () =>
        document.removeEventListener("mousedown", onDoc);
  }, []);

  const toggle = (option) => {
    if (selected.includes(option)) {
      onChange(
          selected.filter((value) => value !== option)
      );
    } else {
      onChange([...selected, option]);
    }
  };

  return (
      <div
          style={{
            position: "relative",
          }}
          ref={ref}
      >
        <button
            onClick={() =>
                !disabled && setOpen((open) => !open)
            }
            disabled={disabled}
            style={{
              width: "100%",
              textAlign: "left",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              padding: "7px 10px",
              borderRadius: 6,
              fontSize: 13,
              fontFamily: sans,
              background: T.surface,
              cursor: disabled ? "default" : "pointer",
              border: `1px solid ${
                  selected.length ? T.amber : T.border
              }`,
              color: selected.length ? T.ink : T.faint,
            }}
        >
        <span
            style={{
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
        >
          {selected.length
              ? `${selected.length} selected: ${selected
                  .slice(0, 2)
                  .join(", ")}${
                  selected.length > 2 ? "…" : ""
              }`
              : placeholder}
        </span>

          <ChevronDown
              size={14}
              color={T.muted}
          />
        </button>

        {open && !disabled && (
            <div
                style={{
                  position: "absolute",
                  zIndex: 50,
                  top: "calc(100% + 4px)",
                  left: 0,
                  right: 0,
                  background: T.surface,
                  border: `1px solid ${T.border}`,
                  borderRadius: 6,
                  maxHeight: 200,
                  overflowY: "auto",
                  boxShadow:
                      "0 4px 14px rgba(0,0,0,0.08)",
                }}
            >
              {options.map((option) => (
                  <label
                      key={option}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        padding: "6px 10px",
                        fontSize: 13,
                        fontFamily: sans,
                        cursor: "pointer",
                      }}
                  >
                    <input
                        type="checkbox"
                        checked={selected.includes(option)}
                        onChange={() => toggle(option)}
                    />

                    <span
                        style={{
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                        }}
                    >
                {option}
              </span>
                  </label>
              ))}

              {options.length === 0 && (
                  <div
                      style={{
                        padding: "8px 10px",
                        fontSize: 12,
                        color: T.faint,
                      }}
                  >
                    No values
                  </div>
              )}
            </div>
        )}
      </div>
  );
}

// ============================================================
// FILTER BUILDER
// ============================================================

function FilterBuilder({
                         columns,
                         filters,
                         onChange,
                         disabled,
                       }) {
  return (
      <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 12,
          }}
      >
        {columns.map((col) => {
          const rule = filters[col.name];
          const active = isRuleActive(rule);

          return (
              <div
                  key={col.name}
                  style={{
                    padding: "10px 12px",
                    borderRadius: 8,
                    border: `1px solid ${
                        active ? T.amber : T.border
                    }`,
                    background: active
                        ? T.amberBg
                        : T.surface,
                  }}
              >
                <div
                    style={{
                      fontSize: 12,
                      fontWeight: 500,
                      color: active
                          ? "#7A4413"
                          : T.muted,
                      marginBottom: 7,
                      display: "flex",
                      justifyContent: "space-between",
                    }}
                >
                  <span>{col.name}</span>

                  <span
                      style={{
                        fontFamily: mono,
                        fontSize: 10,
                        color: T.faint,
                      }}
                  >
                {col.kind}
              </span>
                </div>

                {col.kind === "categorical" ? (
                    <MultiSelect
                        options={col.options}
                        selected={rule.values}
                        disabled={disabled}
                        placeholder="Any value"
                        onChange={(values) =>
                            onChange(col.name, {
                              ...rule,
                              values,
                            })
                        }
                    />
                ) : (
                    <div
                        style={{
                          display: "flex",
                          gap: 6,
                        }}
                    >
                      <select
                          value={rule.operator}
                          disabled={disabled}
                          onChange={(e) =>
                              onChange(col.name, {
                                ...rule,
                                operator: e.target.value,
                              })
                          }
                          style={{
                            padding: "7px 6px",
                            borderRadius: 6,
                            border: `1px solid ${T.border}`,
                            fontSize: 13,
                            fontFamily: sans,
                            background: T.surface,
                            flex: "0 0 auto",
                          }}
                      >
                        <option value="gt">
                          more than
                        </option>
                        <option value="lt">
                          less than
                        </option>
                        <option value="eq">
                          equal to
                        </option>
                        <option value="between">
                          between
                        </option>
                      </select>

                      <input
                          type="number"
                          disabled={disabled}
                          value={rule.value}
                          placeholder={`e.g. ${Math.round(
                              col.min
                          )}`}
                          onChange={(e) =>
                              onChange(col.name, {
                                ...rule,
                                value: e.target.value,
                              })
                          }
                          style={{
                            width: 0,
                            flex: 1,
                            padding: "7px 8px",
                            borderRadius: 6,
                            border: `1px solid ${T.border}`,
                            fontSize: 13,
                            fontFamily: mono,
                          }}
                      />

                      {rule.operator === "between" && (
                          <input
                              type="number"
                              disabled={disabled}
                              value={rule.value2}
                              placeholder={`e.g. ${Math.round(
                                  col.max
                              )}`}
                              onChange={(e) =>
                                  onChange(col.name, {
                                    ...rule,
                                    value2: e.target.value,
                                  })
                              }
                              style={{
                                width: 0,
                                flex: 1,
                                padding: "7px 8px",
                                borderRadius: 6,
                                border: `1px solid ${T.border}`,
                                fontSize: 13,
                                fontFamily: mono,
                              }}
                          />
                      )}
                    </div>
                )}
              </div>
          );
        })}
      </div>
  );
}

// ============================================================
// RESULTS TABLE
// ============================================================

function ResultsTable({
                        columns,
                        rows,
                        matchedIds,
                      }) {
  if (!rows || rows.length === 0) {
    return (
        <div
            style={{
              padding: 24,
              color: T.faint,
              fontSize: 13,
              textAlign: "center",
            }}
        >
          No rows to show.
        </div>
    );
  }

  return (
      <div
          style={{
            overflow: "auto",
            border: `1px solid ${T.border}`,
            borderRadius: 8,
            maxHeight: 460,
          }}
      >
        <table
            style={{
              borderCollapse: "collapse",
              width: "100%",
              fontSize: 12.5,
              fontFamily: sans,
            }}
        >
          <thead>
          <tr
              style={{
                position: "sticky",
                top: 0,
                background: "#EEF0E8",
                zIndex: 1,
              }}
          >
            <th
                style={{
                  padding: "8px 10px",
                  textAlign: "left",
                  borderBottom: `1px solid ${T.border}`,
                  width: 28,
                }}
            />

            {columns.map((c) => (
                <th
                    key={c.name}
                    style={{
                      padding: "8px 10px",
                      textAlign: "left",
                      borderBottom: `1px solid ${T.border}`,
                      color: T.muted,
                      fontWeight: 500,
                      whiteSpace: "nowrap",
                    }}
                >
                  {c.name}
                </th>
            ))}
          </tr>
          </thead>

          <tbody>
          {rows.map((row) => {
            const isMatch =
                matchedIds &&
                matchedIds.has(row._rid);

            return (
                <tr
                    key={row._rid}
                    style={{
                      background: isMatch
                          ? T.amberBg
                          : T.surface,
                      borderLeft: isMatch
                          ? `3px solid ${T.amber}`
                          : "3px solid transparent",
                    }}
                >
                  <td
                      style={{
                        padding: "6px 10px",
                        borderBottom:
                            "1px solid #ECEEE7",
                      }}
                  >
                    {isMatch && (
                        <span
                            title="Matches priority rules"
                            style={{
                              color: T.amber,
                              fontWeight: 700,
                            }}
                        >
                      &#9679;
                    </span>
                    )}
                  </td>

                  {columns.map((c) => (
                      <td
                          key={c.name}
                          style={{
                            padding: "6px 10px",
                            borderBottom:
                                "1px solid #ECEEE7",
                            fontFamily:
                                c.kind === "numeric"
                                    ? mono
                                    : sans,
                            whiteSpace: "nowrap",
                          }}
                      >
                        {row[c.name] === undefined ||
                        row[c.name] === "" ? (
                            <span
                                style={{
                                  color: T.faint,
                                }}
                            >
                        —
                      </span>
                        ) : (
                            String(row[c.name])
                        )}
                      </td>
                  ))}
                </tr>
            );
          })}
          </tbody>
        </table>
      </div>
  );
}

// ============================================================
// COLUMN STATS
// ============================================================

function ColumnCard({
                      col,
                      rows,
                      matchedIds,
                      summed,
                      onToggleSum,
                    }) {
  const [showAll, setShowAll] = useState(false);

  const total = rows.length;

  const entries = useMemo(() => {
    const counts = new Map();

    rows.forEach((r) => {
      const value = r[col.name];

      const key =
          value === "" || value == null
              ? "(empty)"
              : String(value);

      counts.set(
          key,
          (counts.get(key) || 0) + 1
      );
    });

    return [...counts.entries()].sort(
        (a, b) => b[1] - a[1]
    );
  }, [rows, col.name]);

  const visible = showAll
      ? entries
      : entries.slice(0, TOP_N);

  const isNumeric = col.kind === "numeric";

  const hasMatches =
      matchedIds && matchedIds.size > 0;

  return (
      <div
          style={{
            background: T.surface,
            border: `1px solid ${
                summed ? T.teal : T.border
            }`,
            borderRadius: 8,
            padding: "10px 12px",
          }}
      >
        <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 6,
              gap: 8,
            }}
        >
          <div
              style={{
                fontSize: 12.5,
                fontWeight: 500,
                color: T.ink,
                minWidth: 0,
              }}
          >
            <div
                style={{
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
            >
              {col.name}
            </div>

            <div
                style={{
                  fontFamily: mono,
                  fontSize: 10,
                  color: T.faint,
                }}
            >
              {entries.length} distinct values
            </div>
          </div>

          {isNumeric && (
              <label
                  title="Show the sum of all rows for this column"
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 4,
                    fontSize: 12,
                    cursor: "pointer",
                    color: summed
                        ? T.teal
                        : T.muted,
                    fontWeight: summed ? 600 : 400,
                    flexShrink: 0,
                  }}
              >
                <input
                    type="checkbox"
                    checked={summed}
                    onChange={() =>
                        onToggleSum(col.name)
                    }
                />

                <Sigma size={13} />

                Sum
              </label>
          )}
        </div>

        {isNumeric && summed && (
            <div
                style={{
                  background: T.tealBg,
                  borderRadius: 6,
                  padding: "7px 9px",
                  marginBottom: 8,
                  fontSize: 12,
                  color: T.tealDark,
                }}
            >
              <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                  }}
              >
            <span>
              Total, all {total} rows
            </span>

                <span
                    style={{
                      fontFamily: mono,
                      fontWeight: 600,
                    }}
                >
              {fmtNum(
                  sumColumn(rows, col.name)
              )}
            </span>
              </div>

              {hasMatches && (
                  <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        marginTop: 3,
                        color: "#7A4413",
                      }}
                  >
              <span>
                Matched rows only
              </span>

                    <span
                        style={{
                          fontFamily: mono,
                          fontWeight: 600,
                        }}
                    >
                {fmtNum(
                    sumColumn(
                        rows.filter((r) =>
                            matchedIds.has(r._rid)
                        ),
                        col.name
                    )
                )}
              </span>
                  </div>
              )}
            </div>
        )}

        {visible.map(([value, count]) => {
          const pct = total
              ? (count / total) * 100
              : 0;

          return (
              <div
                  key={value}
                  style={{
                    marginBottom: 6,
                  }}
              >
                <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      gap: 8,
                      fontSize: 12,
                    }}
                >
              <span
                  style={{
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                    minWidth: 0,
                  }}
                  title={value}
              >
                {value}
              </span>

                  <span
                      style={{
                        fontFamily: mono,
                        flexShrink: 0,
                        color: T.muted,
                      }}
                  >
                {count} · {pct.toFixed(1)}%
              </span>
                </div>

                <div
                    style={{
                      height: 4,
                      background: T.tealBg,
                      borderRadius: 2,
                      marginTop: 2,
                    }}
                >
                  <div
                      style={{
                        width: `${pct}%`,
                        height: "100%",
                        background: T.teal,
                        borderRadius: 2,
                      }}
                  />
                </div>
              </div>
          );
        })}

        {entries.length > TOP_N && (
            <button
                onClick={() =>
                    setShowAll((s) => !s)
                }
                style={{
                  border: "none",
                  background: "none",
                  padding: 0,
                  cursor: "pointer",
                  fontSize: 11.5,
                  color: T.teal,
                  fontFamily: sans,
                }}
            >
              {showAll
                  ? "Show fewer"
                  : `Show all ${entries.length} values`}
            </button>
        )}
      </div>
  );
}

function StatsPanel({
                      columns,
                      rows,
                      matchedIds,
                      open,
                      onToggle,
                      sumCols,
                      onToggleSum,
                    }) {
  if (!open) {
    return (
        <button
            onClick={onToggle}
            title="Open column stats"
            style={{
              flex: "0 0 auto",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              gap: 8,
              padding: "12px 7px",
              border: `1px solid ${T.border}`,
              borderRadius: 8,
              background: T.surface,
              cursor: "pointer",
              color: T.muted,
              fontFamily: sans,
              fontSize: 12,
            }}
        >
          <ChevronLeft size={14} />
          <BarChart3 size={15} />
          <span
              style={{
                writingMode: "vertical-rl",
              }}
          >
          Column stats
        </span>
        </button>
    );
  }

  return (
      <div
          style={{
            flex: "0 0 300px",
            width: 300,
          }}
      >
        <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 8,
            }}
        >
          <div
              style={{
                fontSize: 13,
                fontWeight: 500,
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
          >
            <BarChart3
                size={14}
                color={T.teal}
            />

            Column stats

            <span
                style={{
                  fontSize: 11,
                  color: T.faint,
                  fontWeight: 400,
                }}
            >
            ({rows.length} rows)
          </span>
          </div>

          <button
              onClick={onToggle}
              title="Close column stats"
              style={{
                border: `1px solid ${T.border}`,
                background: T.surface,
                borderRadius: 6,
                padding: "3px 5px",
                cursor: "pointer",
                display: "flex",
                color: T.muted,
              }}
          >
            <ChevronRight size={14} />
          </button>
        </div>

        <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 8,
              maxHeight: 460,
              overflowY: "auto",
              paddingRight: 2,
            }}
        >
          {rows.length === 0 && (
              <div
                  style={{
                    fontSize: 12,
                    color: T.faint,
                  }}
              >
                No rows to analyze.
              </div>
          )}

          {rows.length > 0 &&
              columns.map((col) => (
                  <ColumnCard
                      key={col.name}
                      col={col}
                      rows={rows}
                      matchedIds={matchedIds}
                      summed={sumCols.has(col.name)}
                      onToggleSum={onToggleSum}
                  />
              ))}
        </div>
      </div>
  );
}

// ============================================================
// SOURCE / WORKING SET SELECTOR
// ============================================================

function SourceSelector({
                          tab,
                          workingSets,
                          originalData,
                          onChangeSource,
                        }) {
  const [joining, setJoining] = useState(
      tab.sourceMode === "join"
  );

  useEffect(() => {
    setJoining(tab.sourceMode === "join");
  }, [tab.sourceMode]);

  const selectedSetIds =
      tab.joinSetIds || [];

  const normalOptions = [
    {
      value: "original",
      label: `Original data (${originalData.length} rows)`,
    },
    ...workingSets.map((set) => ({
      value: `set:${set.id}`,
      label: `${set.name} (${set.rows.length} rows)`,
    })),
  ];

  const addJoinSource = () => {
    const available = workingSets.find(
        (set) =>
            !selectedSetIds.includes(set.id)
    );

    if (!available) return;

    onChangeSource({
      sourceMode: "join",
      joinSetIds: [
        ...selectedSetIds,
        available.id,
      ],
    });
  };

  const updateJoinSet = (index, setId) => {
    const next = [...selectedSetIds];

    next[index] = setId;

    onChangeSource({
      sourceMode: "join",
      joinSetIds: next,
    });
  };

  const removeJoinSet = (index) => {
    const next = selectedSetIds.filter(
        (_, i) => i !== index
    );

    if (next.length === 0) {
      onChangeSource({
        sourceMode: "original",
        joinSetIds: [],
      });
    } else {
      onChangeSource({
        sourceMode: "join",
        joinSetIds: next,
      });
    }
  };

  const switchToNormal = (value) => {
    if (value === "original") {
      onChangeSource({
        sourceMode: "original",
        joinSetIds: [],
      });
      return;
    }

    if (value.startsWith("set:")) {
      onChangeSource({
        sourceMode: "set",
        sourceSetId: value.slice(4),
        joinSetIds: [],
      });
    }
  };

  return (
      <div
          style={{
            marginBottom: 12,
            padding: 12,
            border: `1px solid ${T.border}`,
            borderRadius: 8,
            background: T.surface,
          }}
      >
        <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 7,
              marginBottom: 9,
              fontSize: 12,
              fontWeight: 600,
              color: T.ink,
            }}
        >
          <Database
              size={14}
              color={T.teal}
          />

          Data source
        </div>

        <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 7,
              marginBottom: 10,
              fontSize: 12.5,
            }}
        >
          <input
              type="checkbox"
              checked={joining}
              onChange={(e) => {
                const checked =
                    e.target.checked;

                setJoining(checked);

                if (checked) {
                  const firstSet =
                      workingSets[0];

                  if (firstSet) {
                    onChangeSource({
                      sourceMode: "join",
                      joinSetIds: [
                        firstSet.id,
                      ],
                    });
                  }
                } else {
                  onChangeSource({
                    sourceMode: "original",
                    joinSetIds: [],
                  });
                }
              }}
          />

          <Link2 size={13} color={T.muted} />

          Join working sets
        </div>

        {!joining ? (
            <SelectBox
                value={
                  tab.sourceMode === "original"
                      ? "original"
                      : tab.sourceMode === "set"
                          ? `set:${tab.sourceSetId}`
                          : ""
                }
                onChange={switchToNormal}
                options={normalOptions}
                placeholder="Choose data source"
            />
        ) : (
            <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  gap: 8,
                }}
            >
              {selectedSetIds.map(
                  (selectedId, index) => {
                    const availableOptions =
                        workingSets.map((set) => ({
                          value: set.id,
                          label: `${set.name} (${set.rows.length} rows)`,
                          disabled:
                              selectedSetIds.some(
                                  (id, i) =>
                                      i !== index &&
                                      id === set.id
                              ),
                        }));

                    return (
                        <div
                            key={`${index}-${selectedId}`}
                            style={{
                              display: "flex",
                              gap: 6,
                              alignItems: "center",
                            }}
                        >
                          <div
                              style={{
                                flex: 1,
                              }}
                          >
                            <SelectBox
                                value={selectedId}
                                onChange={(value) =>
                                    updateJoinSet(
                                        index,
                                        value
                                    )
                                }
                                options={
                                  availableOptions
                                }
                                placeholder="Choose working set"
                            />
                          </div>

                          {selectedSetIds.length >
                              1 && (
                                  <button
                                      onClick={() =>
                                          removeJoinSet(
                                              index
                                          )
                                      }
                                      title="Remove this set"
                                      style={{
                                        width: 30,
                                        height: 30,
                                        borderRadius: 6,
                                        border: `1px solid ${T.border}`,
                                        background:
                                        T.surface,
                                        cursor: "pointer",
                                        display: "flex",
                                        alignItems:
                                            "center",
                                        justifyContent:
                                            "center",
                                        color: T.muted,
                                      }}
                                  >
                                    <X size={14} />
                                  </button>
                              )}
                        </div>
                    );
                  }
              )}

              <button
                  onClick={addJoinSource}
                  disabled={
                      selectedSetIds.length >=
                      workingSets.length
                  }
                  style={{
                    border: `1px dashed ${T.borderStrong}`,
                    background: T.bg,
                    color:
                        selectedSetIds.length >=
                        workingSets.length
                            ? T.faint
                            : T.teal,
                    borderRadius: 6,
                    padding: "7px 9px",
                    fontSize: 12,
                    fontFamily: sans,
                    cursor:
                        selectedSetIds.length >=
                        workingSets.length
                            ? "default"
                            : "pointer",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: 5,
                  }}
              >
                <Plus size={13} />

                Add another working set
              </button>

              {workingSets.length === 0 && (
                  <div
                      style={{
                        fontSize: 12,
                        color: T.faint,
                      }}
                  >
                    Save a working set first to
                    join multiple sets.
                  </div>
              )}
            </div>
        )}
      </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================

export default function PriorityFilterApp() {
  const [columns, setColumns] = useState([]);
  const [originalData, setOriginalData] =
      useState([]);

  // ----------------------------------------------------------
  // IMMUTABLE SAVED WORKING SETS
  // ----------------------------------------------------------

  const [workingSets, setWorkingSets] =
      useState([]);

  // ----------------------------------------------------------
  // TABS
  // ----------------------------------------------------------

  const [tabs, setTabs] = useState([]);
  const [activeTabId, setActiveTabId] =
      useState(null);

  // ----------------------------------------------------------
  // UI
  // ----------------------------------------------------------

  const [history, setHistory] =
      useState([]);

  const [historyOpen, setHistoryOpen] =
      useState(false);

  const [statsOpen, setStatsOpen] =
      useState(true);

  const [sumCols, setSumCols] =
      useState(new Set());

  const [fileName, setFileName] =
      useState("");

  const [workingSetsOpen, setWorkingSetsOpen] =
      useState(false);

  const fileInput = useRef(null);

  // ----------------------------------------------------------
  // ACTIVE TAB
  // ----------------------------------------------------------

  const activeTab =
      tabs.find(
          (tab) => tab.id === activeTabId
      ) || null;

  // ==========================================================
  // CREATE NEW TAB
  // ==========================================================

  const createTab = useCallback(
      ({
         sourceMode = "original",
         sourceSetId = null,
         joinSetIds = [],
       } = {}) => {
        const initialData =
            sourceMode === "original"
                ? cloneRows(originalData)
                : sourceMode === "set"
                    ? cloneRows(
                        workingSets.find(
                            (set) =>
                                set.id === sourceSetId
                        )?.rows || []
                    )
                    : combineWorkingSets(
                        workingSets
                            .filter((set) =>
                                joinSetIds.includes(
                                    set.id
                                )
                            )
                            .map((set) => ({
                              ...set,
                              rows: cloneRows(
                                  set.rows
                              ),
                            }))
                    );

        const tab = {
          id: makeTabId(),
          name: `Tab ${tabCounter - 1}`,

          sourceMode,
          sourceSetId,
          joinSetIds,

          // This is the tab's own private copy.
          // It can be changed without changing the
          // source working set.
          data: initialData,

          filters: emptyFilters(
              detectColumns(initialData)
          ),

          result: null,
          matchedIds: new Set(),

          readOnly: false,
        };

        setTabs((prev) => [
          ...prev,
          tab,
        ]);

        setActiveTabId(tab.id);
      },
      [originalData, workingSets]
  );

  // ==========================================================
  // FILE LOADING
  // ==========================================================

  const handleFile = useCallback(
      (file) => {
        const reader = new FileReader();

        reader.onload = (e) => {
          const wb = XLSX.read(
              e.target.result,
              {
                type: "binary",
              }
          );

          const sheet =
              wb.Sheets[wb.SheetNames[0]];

          const rows =
              XLSX.utils.sheet_to_json(sheet, {
                defval: "",
              });

          const cleanRows = rows.map(
              (row, index) => ({
                ...row,
                _rid: `original-${Date.now()}-${index}`,
              })
          );

          const cols =
              detectColumns(cleanRows);

          setColumns(cols);
          setOriginalData(cleanRows);
          setWorkingSets([]);
          setFileName(file.name);
          setSumCols(new Set());
          setHistory([]);

          const tab = {
            id: makeTabId(),
            name: "Tab 1",

            sourceMode: "original",
            sourceSetId: null,
            joinSetIds: [],

            data: cloneRows(cleanRows),

            filters: emptyFilters(cols),

            result: null,
            matchedIds: new Set(),

            readOnly: false,
          };

          setTabs([tab]);
          setActiveTabId(tab.id);
        };

        reader.readAsBinaryString(file);
      },
      []
  );

  // ==========================================================
  // TAB UPDATE
  // ==========================================================

  const updateTab = (
      id,
      patch
  ) => {
    setTabs((prev) =>
        prev.map((tab) =>
            tab.id === id
                ? {
                  ...tab,
                  ...patch,
                }
                : tab
        )
    );
  };

  // ==========================================================
  // TAB NAME
  // ==========================================================

  const renameTab = (tab) => {
    const newName = window.prompt(
        "Enter a name for this tab:",
        tab.name
    );

    if (
        newName === null ||
        !newName.trim()
    ) {
      return;
    }

    updateTab(tab.id, {
      name: newName.trim(),
    });
  };

  // ==========================================================
  // FILTER RULE
  // ==========================================================

  const setRule = (
      colName,
      rule
  ) => {
    if (
        !activeTab ||
        activeTab.readOnly
    ) {
      return;
    }

    updateTab(activeTab.id, {
      filters: {
        ...activeTab.filters,
        [colName]: rule,
      },
    });
  };

  // ==========================================================
  // CHANGE DATA SOURCE
  // ==========================================================

  const changeTabSource = ({
                             sourceMode,
                             sourceSetId = null,
                             joinSetIds = [],
                           }) => {
    if (!activeTab) return;

    let newData = [];

    if (sourceMode === "original") {
      newData = cloneRows(
          originalData
      );
    }

    if (sourceMode === "set") {
      const selected =
          workingSets.find(
              (set) =>
                  set.id === sourceSetId
          );

      newData = cloneRows(
          selected?.rows || []
      );
    }

    if (sourceMode === "join") {
      const selectedSets =
          workingSets
              .filter((set) =>
                  joinSetIds.includes(
                      set.id
                  )
              )
              .map((set) => ({
                ...set,
                rows: cloneRows(
                    set.rows
                ),
              }));

      newData =
          combineWorkingSets(
              selectedSets
          );
    }

    const newColumns =
        detectColumns(newData);

    updateTab(activeTab.id, {
      sourceMode,
      sourceSetId,
      joinSetIds,

      data: newData,

      filters:
          emptyFilters(newColumns),

      result: null,
      matchedIds: new Set(),
    });
  };

  // ==========================================================
  // APPLY FILTER
  // ==========================================================

  const applyTab = () => {
    if (!activeTab) return;

    const { result, matchedIds } =
        applyPriority(
            activeTab.data,
            activeTab.filters
        );

    updateTab(activeTab.id, {
      result,
      matchedIds,
    });
  };

  // ==========================================================
  // KEEP ONLY MATCHED
  // ==========================================================

  const keepOnlyMatchedInTab = () => {
    if (
        !activeTab ||
        !activeTab.result ||
        activeTab.matchedIds.size === 0
    ) {
      return;
    }

    const nextData =
        activeTab.data.filter((row) =>
            activeTab.matchedIds.has(
                row._rid
            )
        );

    const nextColumns =
        detectColumns(nextData);

    updateTab(activeTab.id, {
      data: nextData,

      filters:
          emptyFilters(nextColumns),

      result: null,
      matchedIds: new Set(),
    });
  };

  // ==========================================================
  // REMOVE MATCHED
  // ==========================================================

  const removeMatchedFromTab = () => {
    if (
        !activeTab ||
        !activeTab.result ||
        activeTab.matchedIds.size === 0
    ) {
      return;
    }

    const nextData =
        activeTab.data.filter(
            (row) =>
                !activeTab.matchedIds.has(
                    row._rid
                )
        );

    const nextColumns =
        detectColumns(nextData);

    updateTab(activeTab.id, {
      data: nextData,

      filters:
          emptyFilters(nextColumns),

      result: null,
      matchedIds: new Set(),
    });
  };

  // ==========================================================
  // SAVE CURRENT TAB AS WORKING SET
  // ==========================================================

  const saveCurrentAsWorkingSet =
      () => {
        if (!activeTab) return;

        const dataToSave =
            activeTab.result
                ? activeTab.result
                : activeTab.data;

        if (!dataToSave.length) {
          window.alert(
              "There are no rows to save."
          );
          return;
        }

        const defaultName =
            `${activeTab.name} — ${
                workingSets.length + 1
            }`;

        const name = window.prompt(
            "Name this working set:",
            defaultName
        );

        if (
            name === null ||
            !name.trim()
        ) {
          return;
        }

        const newSet = {
          id: makeWorkingSetId(),

          name: name.trim(),

          createdAt:
              new Date().toLocaleString(),

          sourceTabName:
          activeTab.name,

          rows: cloneRows(
              dataToSave
          ),
        };

        setWorkingSets((prev) => [
          ...prev,
          newSet,
        ]);

        // After saving, clear the result state
        // so the tab continues working on its
        // private dataset.
        updateTab(activeTab.id, {
          data: cloneRows(
              dataToSave
          ),
          result: null,
          matchedIds: new Set(),
          filters: emptyFilters(
              detectColumns(
                  dataToSave
              )
          ),
        });
      };

  // ==========================================================
  // DELETE WORKING SET
  // ==========================================================

  const deleteWorkingSet = (
      setId
  ) => {
    const set =
        workingSets.find(
            (item) =>
                item.id === setId
        );

    if (!set) return;

    const confirmed =
        window.confirm(
            `Delete working set "${set.name}"?\n\nThis does not affect any tabs that are already working from it.`
        );

    if (!confirmed) return;

    setWorkingSets((prev) =>
        prev.filter(
            (item) =>
                item.id !== setId
        )
    );
  };

  // ==========================================================
  // DOWNLOAD CURRENT TAB
  // ==========================================================

  const downloadTabExcel = () => {
    if (!activeTab) return;

    const data =
        activeTab.result ||
        activeTab.data;

    if (!data.length) {
      window.alert(
          "There are no rows to download."
      );
      return;
    }

    const exportData =
        data.map((row) => {
          const clean = {};

          Object.keys(row).forEach(
              (key) => {
                if (
                    key !== "_rid" &&
                    key !== "_sourceSet"
                ) {
                  clean[key] =
                      row[key];
                }
              }
          );

          return clean;
        });

    const worksheet =
        XLSX.utils.json_to_sheet(
            exportData
        );

    const workbook =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Data"
    );

    const safeName =
        activeTab.name
            .replace(
                /[^a-z0-9-_ ]/gi,
                ""
            )
            .trim() ||
        "working-set";

    XLSX.writeFile(
        workbook,
        `${safeName}.xlsx`
    );
  };

  // ==========================================================
  // SAVE TO HISTORY
  // ==========================================================

  const saveToHistory = () => {
    if (
        !activeTab ||
        !activeTab.result
    ) {
      return;
    }

    const snap = {
      id: makeHistoryId(),

      name: activeTab.name,

      timestamp:
          new Date().toLocaleString(),

      columns:
          detectColumns(
              activeTab.result
          ),

      filters:
      activeTab.filters,

      result: cloneRows(
          activeTab.result
      ),

      matchedIds:
          new Set(
              activeTab.matchedIds
          ),

      matchedCount:
      activeTab.matchedIds.size,

      totalCount:
      activeTab.result.length,
    };

    setHistory((prev) => [
      snap,
      ...prev,
    ]);
  };

  // ==========================================================
  // OPEN HISTORY
  // ==========================================================

  const openHistoryTab = (
      snap
  ) => {
    const data =
        cloneRows(
            snap.result
        );

    const t = {
      id: `histview-${snap.id}`,

      name: snap.name,

      sourceMode: "history",
      sourceSetId: null,
      joinSetIds: [],

      data,

      filters: snap.filters,

      result: data,

      matchedIds:
          new Set(
              snap.matchedIds
          ),

      readOnly: true,

      historyMeta: snap,
    };

    setTabs((prev) => [
      ...prev.filter(
          (tab) =>
              tab.id !== t.id
      ),
      t,
    ]);

    setActiveTabId(t.id);
    setHistoryOpen(false);
  };

  // ==========================================================
  // CLOSE TAB
  // ==========================================================

  const closeTab = (id) => {
    setTabs((prev) => {
      const next =
          prev.filter(
              (tab) =>
                  tab.id !== id
          );

      if (
          activeTabId === id
      ) {
        if (next.length) {
          setActiveTabId(
              next[next.length - 1]
                  .id
          );
        } else {
          setActiveTabId(null);
        }
      }

      return next;
    });
  };

  // ==========================================================
  // TOGGLE SUM
  // ==========================================================

  const toggleSum = (
      colName
  ) => {
    setSumCols((prev) => {
      const next =
          new Set(prev);

      if (
          next.has(colName)
      ) {
        next.delete(colName);
      } else {
        next.add(colName);
      }

      return next;
    });
  };

  // ==========================================================
  // RESET FILTERS
  // ==========================================================

  const clearRules = () => {
    if (!activeTab) return;

    updateTab(activeTab.id, {
      filters: emptyFilters(
          detectColumns(
              activeTab.data
          )
      ),
      result: null,
      matchedIds: new Set(),
    });
  };

  // ==========================================================
  // OPEN NEW NORMAL TAB
  // ==========================================================

  const openNewTab = () => {
    createTab({
      sourceMode: "original",
    });
  };

  // ==========================================================
  // CURRENT DISPLAY
  // ==========================================================

  const shownRows = activeTab
      ? activeTab.result ||
      activeTab.data
      : [];

  const currentColumns =
      detectColumns(shownRows);

  const currentDataColumns =
      detectColumns(
          activeTab?.data || []
      );

  // ==========================================================
  // EMPTY STATE
  // ==========================================================

  if (!columns.length) {
    return (
        <div
            style={{
              fontFamily: sans,
              background: T.bg,
              minHeight: 420,
              borderRadius: 12,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: 40,
            }}
        >
          <div
              style={{
                textAlign: "center",
                maxWidth: 400,
              }}
          >
            <FileSpreadsheet
                size={30}
                color={T.teal}
                style={{
                  marginBottom: 12,
                }}
            />

            <div
                style={{
                  fontSize: 16,
                  color: T.ink,
                  marginBottom: 4,
                }}
            >
              Load a spreadsheet to begin
            </div>

            <div
                style={{
                  fontSize: 13,
                  color: T.muted,
                  marginBottom: 18,
                }}
            >
              Every column becomes a
              priority control automatically.
              Saved working sets can be used
              independently in new tabs or
              combined into a new dataset.
            </div>

            <Btn
                variant="primary"
                icon={Upload}
                onClick={() =>
                    fileInput.current?.click()
                }
            >
              Choose .xlsx file
            </Btn>

            <input
                ref={fileInput}
                type="file"
                accept=".xlsx,.xls"
                style={{
                  display: "none",
                }}
                onChange={(e) => {
                  if (
                      e.target.files[0]
                  ) {
                    handleFile(
                        e.target.files[0]
                    );
                  }
                }}
            />
          </div>
        </div>
    );
  }

  // ==========================================================
  // MAIN UI
  // ==========================================================

  return (
      <div
          style={{
            fontFamily: sans,
            background: T.bg,
            borderRadius: 12,
            padding: 16,
            color: T.ink,
          }}
      >
        {/* ======================================================
          HEADER
      ====================================================== */}

        <div
            style={{
              display: "flex",
              justifyContent:
                  "space-between",
              alignItems: "center",
              marginBottom: 12,
              gap: 12,
            }}
        >
          <div>
            <div
                style={{
                  fontSize: 15,
                  fontWeight: 500,
                }}
            >
              {fileName}
            </div>

            <div
                style={{
                  fontSize: 12,
                  color: T.muted,
                }}
            >
              {originalData.length} original
              rows · {workingSets.length} saved
              working set
              {workingSets.length !== 1
                  ? "s"
                  : ""}
            </div>
          </div>

          <div
              style={{
                display: "flex",
                gap: 8,
                flexWrap: "wrap",
                justifyContent:
                    "flex-end",
              }}
          >
            <Btn
                icon={Upload}
                onClick={() =>
                    fileInput.current?.click()
                }
            >
              Load new file
            </Btn>

            <input
                ref={fileInput}
                type="file"
                accept=".xlsx,.xls"
                style={{
                  display: "none",
                }}
                onChange={(e) => {
                  if (
                      e.target.files[0]
                  ) {
                    handleFile(
                        e.target.files[0]
                    );
                  }
                }}
            />

            <Btn
                icon={Database}
                onClick={() =>
                    setWorkingSetsOpen(
                        (open) => !open
                    )
                }
                variant={
                  workingSetsOpen
                      ? "primary"
                      : "ghost"
                }
            >
              Working sets ({workingSets.length})
            </Btn>

            <Btn
                icon={History}
                onClick={() =>
                    setHistoryOpen(
                        (open) => !open
                    )
                }
                variant={
                  historyOpen
                      ? "primary"
                      : "ghost"
                }
            >
              History ({history.length})
            </Btn>
          </div>
        </div>

        {/* ======================================================
          WORKING SET MANAGER
      ====================================================== */}

        {workingSetsOpen && (
            <div
                style={{
                  marginBottom: 12,
                  border: `1px solid ${T.border}`,
                  borderRadius: 8,
                  background: T.surface,
                  padding: 12,
                }}
            >
              <div
                  style={{
                    display: "flex",
                    justifyContent:
                        "space-between",
                    alignItems: "center",
                    marginBottom: 10,
                  }}
              >
                <div
                    style={{
                      fontSize: 13,
                      fontWeight: 600,
                      display: "flex",
                      alignItems:
                          "center",
                      gap: 6,
                    }}
                >
                  <Database
                      size={14}
                      color={T.teal}
                  />

                  Saved working sets
                </div>

                <div
                    style={{
                      fontSize: 11,
                      color: T.faint,
                    }}
                >
                  Saved sets are immutable
                  sources.
                </div>
              </div>

              {workingSets.length === 0 && (
                  <div
                      style={{
                        padding: 10,
                        background: T.bg,
                        borderRadius: 6,
                        fontSize: 12,
                        color: T.faint,
                      }}
                  >
                    No working sets saved yet.
                    Use "Save as working set" on
                    a tab after filtering your
                    data.
                  </div>
              )}

              {workingSets.map((set) => (
                  <div
                      key={set.id}
                      style={{
                        display: "flex",
                        alignItems:
                            "center",
                        justifyContent:
                            "space-between",
                        gap: 10,
                        padding:
                            "9px 10px",
                        borderBottom:
                            `1px solid #EEF0E8`,
                      }}
                  >
                    <div
                        style={{
                          minWidth: 0,
                        }}
                    >
                      <div
                          style={{
                            fontSize: 13,
                            fontWeight: 500,
                          }}
                      >
                        {set.name}
                      </div>

                      <div
                          style={{
                            fontSize: 11,
                            color: T.faint,
                            marginTop: 2,
                          }}
                      >
                        {set.rows.length} rows
                        {" · "}
                        saved {set.createdAt}
                        {" · "}
                        from {set.sourceTabName}
                      </div>
                    </div>

                    <div
                        style={{
                          display: "flex",
                          gap: 6,
                        }}
                    >
                      <Btn
                          icon={Trash2}
                          variant="danger"
                          onClick={() =>
                              deleteWorkingSet(
                                  set.id
                              )
                          }
                      >
                        Delete
                      </Btn>
                    </div>
                  </div>
              ))}
            </div>
        )}

        {/* ======================================================
          HISTORY
      ====================================================== */}

        {historyOpen && (
            <div
                style={{
                  marginBottom: 12,
                  border: `1px solid ${T.border}`,
                  borderRadius: 8,
                  background: T.surface,
                  padding: 12,
                  maxHeight: 220,
                  overflowY: "auto",
                }}
            >
              {history.length === 0 && (
                  <div
                      style={{
                        fontSize: 13,
                        color: T.faint,
                      }}
                  >
                    No saved runs yet. Apply a
                    priority and click "Save to
                    history".
                  </div>
              )}

              {history.map((h) => (
                  <div
                      key={h.id}
                      onClick={() =>
                          openHistoryTab(h)
                      }
                      style={{
                        display: "flex",
                        justifyContent:
                            "space-between",
                        alignItems: "center",
                        padding:
                            "8px 10px",
                        borderRadius: 6,
                        cursor: "pointer",
                        fontSize: 13,
                        borderBottom:
                            "1px solid #EEF0E8",
                      }}
                  >
                    <div
                        style={{
                          display: "flex",
                          alignItems:
                              "center",
                          gap: 8,
                        }}
                    >
                      <Clock
                          size={13}
                          color={T.muted}
                      />

                      <span
                          style={{
                            fontWeight: 500,
                          }}
                      >
                  {h.name}
                </span>

                      <span
                          style={{
                            color: T.faint,
                            fontSize: 12,
                          }}
                      >
                  {h.timestamp}
                </span>
                    </div>

                    <span
                        style={{
                          color: T.amber,
                          fontSize: 12,
                          fontFamily: mono,
                        }}
                    >
                {h.matchedCount} matched
                      {" / "}
                      {h.totalCount} rows
              </span>
                  </div>
              ))}
            </div>
        )}

        {/* ======================================================
          TAB BAR
      ====================================================== */}

        <div
            style={{
              display: "flex",
              alignItems: "flex-end",
              gap: 2,
              borderBottom: `1px solid ${T.border}`,
              marginBottom: 14,
              overflowX: "auto",
            }}
        >
          {tabs.map((tab) => {
            const active =
                tab.id === activeTabId;

            return (
                <div
                    key={tab.id}
                    onClick={() =>
                        setActiveTabId(
                            tab.id
                        )
                    }
                    style={{
                      display: "flex",
                      alignItems:
                          "center",
                      gap: 6,
                      padding:
                          "8px 12px",
                      borderRadius:
                          "8px 8px 0 0",
                      cursor: "pointer",
                      fontSize: 13,
                      background: active
                          ? T.surface
                          : "transparent",
                      border: active
                          ? `1px solid ${T.border}`
                          : "1px solid transparent",
                      borderBottom:
                          active
                              ? `1px solid ${T.surface}`
                              : "none",
                      marginBottom: -1,
                      color: active
                          ? T.ink
                          : T.muted,
                      borderStyle:
                          tab.readOnly
                              ? "dashed"
                              : "solid",
                      whiteSpace:
                          "nowrap",
                    }}
                >
                  {tab.readOnly && (
                      <Lock
                          size={11}
                          color={T.faint}
                      />
                  )}

                  <span>
                {tab.name}
              </span>

                  <button
                      onClick={(e) => {
                        e.stopPropagation();
                        renameTab(tab);
                      }}
                      title="Rename tab"
                      style={{
                        border: "none",
                        background:
                            "transparent",
                        padding: 0,
                        cursor: "pointer",
                        color: T.faint,
                        display: "flex",
                      }}
                  >
                    <Pencil size={11} />
                  </button>

                  <span
                      onClick={(e) => {
                        e.stopPropagation();
                        closeTab(
                            tab.id
                        );
                      }}
                      style={{
                        color: T.faint,
                        display: "flex",
                      }}
                  >
                <X size={12} />
              </span>
                </div>
            );
          })}

          <button
              onClick={openNewTab}
              style={{
                border: "none",
                background: "none",
                cursor: "pointer",
                padding:
                    "8px 10px",
                color: T.muted,
                display: "flex",
                alignItems:
                    "center",
              }}
              title="New tab"
          >
            <Plus size={15} />
          </button>
        </div>

        {/* ======================================================
          ACTIVE TAB
      ====================================================== */}

        {activeTab && (
            <div
                style={{
                  display: "grid",
                  gridTemplateColumns:
                      "300px minmax(0, 1fr)",
                  gap: 16,
                }}
            >
              {/* ==================================================
              LEFT SIDE
          ================================================== */}

              <div>
                {activeTab.readOnly &&
                    activeTab.historyMeta && (
                        <div
                            style={{
                              fontSize: 12,
                              color: T.muted,
                              background:
                              T.tealBg,
                              border: `1px solid ${T.border}`,
                              borderRadius: 6,
                              padding:
                                  "8px 10px",
                              marginBottom: 10,
                            }}
                        >
                          Read-only view of a
                          saved run from{" "}
                          {
                            activeTab
                                .historyMeta
                                .timestamp
                          }
                          .
                        </div>
                    )}

                {!activeTab.readOnly && (
                    <>
                      {/* ------------------------------------------
                    DATA SOURCE
                ------------------------------------------ */}

                      <SourceSelector
                          tab={activeTab}
                          workingSets={
                            workingSets
                          }
                          originalData={
                            originalData
                          }
                          onChangeSource={
                            changeTabSource
                          }
                      />

                      {/* ------------------------------------------
                    TAB DATA INFO
                ------------------------------------------ */}

                      <div
                          style={{
                            padding:
                                "8px 10px",
                            marginBottom: 10,
                            background:
                            T.blueBg,
                            border: `1px solid ${T.border}`,
                            borderRadius: 6,
                            fontSize: 12,
                            color: T.muted,
                          }}
                      >
                        <div
                            style={{
                              fontWeight: 600,
                              color: T.ink,
                              marginBottom: 3,
                            }}
                        >
                          This tab has its own
                          working copy
                        </div>

                        <div>
                          {activeTab.data.length}{" "}
                          rows are currently in
                          this tab. Changes here do
                          not modify the saved
                          working set you selected.
                        </div>
                      </div>

                      {/* ------------------------------------------
                    FILTERS
                ------------------------------------------ */}

                      <FilterBuilder
                          columns={
                            currentDataColumns
                          }
                          filters={
                            activeTab.filters
                          }
                          onChange={setRule}
                          disabled={
                            activeTab.readOnly
                          }
                      />
                    </>
                )}

                {activeTab.readOnly && (
                    <FilterBuilder
                        columns={
                          currentDataColumns
                        }
                        filters={
                          activeTab.filters
                        }
                        onChange={setRule}
                        disabled={true}
                    />
                )}
              </div>

              {/* ==================================================
              RIGHT SIDE
          ================================================== */}

              <div
                  style={{
                    minWidth: 0,
                  }}
              >
                {!activeTab.readOnly && (
                    <div
                        style={{
                          display: "flex",
                          gap: 8,
                          marginBottom: 10,
                          flexWrap: "wrap",
                        }}
                    >
                      <Btn
                          variant="primary"
                          icon={
                            ArrowDownWideNarrow
                          }
                          onClick={applyTab}
                      >
                        Apply priority
                      </Btn>

                      <Btn
                          icon={Save}
                          variant="primary"
                          onClick={
                            saveCurrentAsWorkingSet
                          }
                      >
                        Save as working set
                      </Btn>

                      <Btn
                          icon={Download}
                          onClick={
                            downloadTabExcel
                          }
                          disabled={
                            !activeTab.data.length
                          }
                      >
                        Download Excel
                      </Btn>

                      <Btn
                          icon={Check}
                          disabled={
                            !activeTab.result
                          }
                          onClick={
                            saveToHistory
                          }
                      >
                        Save to history
                      </Btn>

                      <Btn
                          icon={Check}
                          variant="primary"
                          disabled={
                              !activeTab.result ||
                              activeTab.matchedIds
                                  .size === 0
                          }
                          onClick={
                            keepOnlyMatchedInTab
                          }
                      >
                        Keep only matched
                      </Btn>

                      <Btn
                          icon={Trash2}
                          variant="danger"
                          disabled={
                              !activeTab.result ||
                              activeTab.matchedIds
                                  .size === 0
                          }
                          onClick={
                            removeMatchedFromTab
                          }
                      >
                        Remove matched
                      </Btn>

                      <Btn
                          icon={RotateCcw}
                          onClick={clearRules}
                      >
                        Clear rules
                      </Btn>
                    </div>
                )}

                {/* ----------------------------------------------
                RESULT MESSAGE
            ---------------------------------------------- */}

                {activeTab.result && (
                    <div
                        style={{
                          fontSize: 12,
                          color: T.muted,
                          marginBottom: 8,
                          display: "flex",
                          gap: 8,
                          alignItems: "center",
                        }}
                    >
                <span
                    style={{
                      color: T.amber,
                      fontWeight: 500,
                    }}
                >
                  {
                    activeTab
                        .matchedIds
                        .size
                  }{" "}
                  rows
                </span>

                      match the priority
                      rules and are pinned to
                      the top.
                    </div>
                )}

                {/* ----------------------------------------------
                TABLE + STATS
            ---------------------------------------------- */}

                <div
                    style={{
                      display: "flex",
                      gap: 12,
                      alignItems:
                          "flex-start",
                    }}
                >
                  <div
                      style={{
                        flex: 1,
                        minWidth: 0,
                      }}
                  >
                    <ResultsTable
                        columns={
                          currentColumns
                        }
                        rows={shownRows}
                        matchedIds={
                          activeTab.matchedIds
                        }
                    />
                  </div>

                  <StatsPanel
                      columns={
                        currentColumns
                      }
                      rows={shownRows}
                      matchedIds={
                        activeTab.matchedIds
                      }
                      open={statsOpen}
                      onToggle={() =>
                          setStatsOpen(
                              (open) => !open
                          )
                      }
                      sumCols={sumCols}
                      onToggleSum={
                        toggleSum
                      }
                  />
                </div>
              </div>
            </div>
        )}
      </div>
  );
}
