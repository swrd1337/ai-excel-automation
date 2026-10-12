/*
 * The v1 tool catalog.
 *
 * One definition per tool is the single source of truth for BOTH consumers:
 *   - the JSON Schema handed to the model, and
 *   - the validator that guards the arguments the model hands back.
 *
 * Execution lives in the taskpane (addin/executor) because only code running
 * inside Excel can call Office.js. The taskpane announces which tools it can
 * execute (`capabilities`), and the server intersects that with this list — so a
 * taskpane that hasn't been reloaded yet can never be sent a tool it can't run.
 *
 * `mutates` drives snapshotting and approval in the safety layer.
 *
 * `executesOn` decides who runs the tool:
 *   'taskpane' (default) — needs Office.js, so the taskpane executes it
 *   'server'             — needs the network or a secret, so the server executes it inline
 * Server tools never cross the wire: the server resolves them itself and only
 * hands taskpane tools to the client.
 */

import type { ApiTool } from './ai.js';

export type JsonSchemaType = 'object' | 'array' | 'string' | 'number' | 'integer' | 'boolean';

export interface JsonSchema {
  type?: JsonSchemaType;
  description?: string;
  properties?: Record<string, JsonSchema>;
  required?: string[];
  items?: JsonSchema;
  enum?: string[];
  additionalProperties?: boolean;
}

export interface ToolDefinition {
  name: string;
  description: string;
  executesOn?: 'taskpane' | 'server';
  mutates?: boolean;
  parameters: JsonSchema;
}

/** Minimal JSON-Schema validator — supports exactly what we use, no dependency. */
export function validateArgs(schema: JsonSchema, value: unknown, path = ''): string[] {
  const errors: string[] = [];
  const where = path || 'arguments';

  if (schema.type === 'object') {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
      return [`${where} must be an object`];
    }
    const record = value as Record<string, unknown>;

    for (const key of schema.required || []) {
      if (record[key] === undefined || record[key] === null) errors.push(`${key} is required`);
    }
    for (const [key, sub] of Object.entries(schema.properties || {})) {
      if (record[key] === undefined) continue;
      errors.push(...validateArgs(sub, record[key], path ? `${path}.${key}` : key));
    }
    if (schema.additionalProperties === false) {
      for (const key of Object.keys(record)) {
        if (!(schema.properties && key in schema.properties)) errors.push(`${key} is not a recognised argument`);
      }
    }
    return errors;
  }

  if (schema.type === 'array') {
    if (!Array.isArray(value)) return [`${where} must be an array`];
    if (schema.items) {
      value.forEach((item, i) => errors.push(...validateArgs(schema.items as JsonSchema, item, `${where}[${i}]`)));
    }
    return errors;
  }

  const actual = value === null ? 'null' : Array.isArray(value) ? 'array' : typeof value;
  if (schema.type === 'integer' && !(actual === 'number' && Number.isInteger(value))) {
    return [`${where} must be an integer, got ${actual}`];
  }
  if (schema.type === 'number' && actual !== 'number') return [`${where} must be a number, got ${actual}`];
  if (schema.type === 'string' && actual !== 'string') return [`${where} must be a string, got ${actual}`];
  if (schema.type === 'boolean' && actual !== 'boolean') return [`${where} must be a boolean, got ${actual}`];

  if (schema.enum && !schema.enum.includes(value as string)) {
    errors.push(`${where} must be one of ${schema.enum.join(', ')}`);
  }
  return errors;
}

const SHEET: JsonSchema = {
  type: 'string',
  description: 'Worksheet name. Defaults to the active sheet when omitted.',
};

export const TOOLS: ToolDefinition[] = [
  {
    name: 'get_workbook_overview',
    description:
      'List every worksheet with its used range and dimensions. Call this first when you do not ' +
      'already know the shape of the workbook.',
    mutates: false,
    parameters: { type: 'object', properties: {}, required: [], additionalProperties: false },
  },
  {
    name: 'read_range',
    description:
      'Read cell values from a range using A1 notation, e.g. "A1:C10". Returns a 2D array of rows. ' +
      'Always read a range before rewriting it.',
    mutates: false,
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'A1-notation range, e.g. "B2:D20".' },
        sheet: SHEET,
      },
      required: ['address'],
      additionalProperties: false,
    },
  },
  {
    name: 'get_selection',
    description: 'Get the range the user currently has selected, with its values and formulas.',
    mutates: false,
    parameters: { type: 'object', properties: {}, required: [], additionalProperties: false },
  },
  {
    name: 'write_range',
    description:
      'Write a 2D array of values into a range. The array is rows-of-columns and must match the ' +
      'range shape. Use null for an empty cell.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'Top-left anchor or exact range, e.g. "A1" or "A1:C3".' },
        values: {
          type: 'array',
          description: 'Rows of values, e.g. [["Name","Qty"],["Widget",4]].',
          items: { type: 'array' },
        },
        sheet: SHEET,
      },
      required: ['address', 'values'],
      additionalProperties: false,
    },
  },
  {
    name: 'set_formula',
    description:
      'Write formulas into a range. Formulas must start with "=" and use the sheet\'s own locale ' +
      'separators (comma). Example: [["=SUM(B2:B10)"]]',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'Range to receive the formulas, e.g. "B11".' },
        formulas: {
          type: 'array',
          description: 'Rows of formula strings, e.g. [["=SUM(B2:B10)"]]',
          items: { type: 'array' },
        },
        sheet: SHEET,
      },
      required: ['address', 'formulas'],
      additionalProperties: false,
    },
  },
  {
    name: 'format_range',
    description:
      'Apply formatting to a range: bold, italic, font colour, fill colour, number format, ' +
      'and horizontal alignment. Only the properties you pass are changed.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'A1-notation range, e.g. "A1:C1".' },
        sheet: SHEET,
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        fontColor: { type: 'string', description: 'Hex colour, e.g. "#FFFFFF".' },
        fillColor: { type: 'string', description: 'Hex colour, e.g. "#1F6FEB".' },
        numberFormat: { type: 'string', description: 'Excel number format code, e.g. "0.00%" or "#,##0.00".' },
        horizontalAlignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
      },
      required: ['address'],
      additionalProperties: false,
    },
  },
  {
    name: 'sort_range',
    description:
      'Sort a range by one or more columns. Values are rearranged inside the range only. ' +
      'Give column offsets relative to the first column of the range (0 = the first column).',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'The range to sort, e.g. "A1:C20".' },
        sheet: SHEET,
        keys: {
          type: 'array',
          description: 'Sort keys in priority order, e.g. [{column:2, ascending:false}]',
          items: {
            type: 'object',
            properties: {
              column: { type: 'integer', description: 'Zero-based offset from the range first column.' },
              ascending: { type: 'boolean', description: 'Default true.' },
            },
            required: ['column'],
            additionalProperties: false,
          },
        },
        hasHeaders: { type: 'boolean', description: 'Whether the first row is a header. Default true.' },
      },
      required: ['address', 'keys'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_table',
    description:
      'Turn an existing range into an Excel table (what Excel calls a ListObject): it gains filter ' +
      'buttons, banded rows and a name you can refer to in formulas. The cells must already hold ' +
      'the data — a table does not move, sort or add anything. You must choose the table name.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description:
            'The full range the table covers, including its header row, e.g. "A1:D20". A single ' +
            'cell is refused because Excel would expand it unpredictably.',
        },
        name: {
          type: 'string',
          description:
            'A unique name for the table, e.g. "SalesTable". Letters, digits and underscores only, ' +
            'starting with a letter or underscore.',
        },
        sheet: SHEET,
        hasHeaders: { type: 'boolean', description: 'Whether the first row is a header row. Default true.' },
      },
      required: ['address', 'name'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_chart',
    description:
      'Create a chart from a range. The range must already hold the data, with its header row or ' +
      'column first — a chart does not add or move anything. You must choose the chart name.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'The range the chart is built from, including headers, e.g. "A1:C13".',
        },
        name: {
          type: 'string',
          description: 'A unique name for the chart on this sheet, e.g. "SalesChart". Spaces are allowed.',
        },
        type: {
          type: 'string',
          description: 'Chart type. Defaults to ColumnClustered.',
          enum: [
            'ColumnClustered', 'ColumnStacked', 'BarClustered', 'BarStacked',
            'Line', 'LineMarkers', 'Area', 'Pie', 'Doughnut',
            'XYScatter', 'XYScatterLines', 'Radar', 'Histogram',
            'Waterfall', 'Funnel', 'Treemap', 'Sunburst',
          ],
        },
        seriesBy: {
          type: 'string',
          description:
            'Whether the series come from the columns or the rows of the range. Defaults to Columns, ' +
            'which is what a table with headers across the top wants.',
          enum: ['Auto', 'Columns', 'Rows'],
        },
        position: {
          type: 'string',
          description: 'Top-left cell to float the chart over, e.g. "E2". Omit to let Excel place it.',
        },
        sheet: SHEET,
      },
      required: ['address', 'name'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_pivot_table',
    description:
      'Summarise a range with a pivot table. Fields are named by the *source headers*, not by column ' +
      'letters: rows are the fields to group by, values are the fields to total. You must choose the ' +
      'pivot table name. It writes computed cells into the destination, so pick an empty area.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'A unique name for the pivot table on this sheet, e.g. "SalesByRegion".',
        },
        source: {
          type: 'string',
          description: 'The range to summarise, including its header row, e.g. "A1:C100".',
        },
        destination: {
          type: 'string',
          description: 'Top-left cell for the pivot, e.g. "E2". Must be empty space.',
        },
        rows: {
          type: 'array',
          description: 'Source header names to group rows by, e.g. ["Region"].',
          items: { type: 'string' },
        },
        values: {
          type: 'array',
          description: 'Source header names to total, e.g. ["Units"].',
          items: { type: 'string' },
        },
        sheet: SHEET,
      },
      required: ['name', 'source', 'destination'],
      additionalProperties: false,
    },
  },
  {
    name: 'remove_duplicates',
    description:
      'Delete rows of a range that duplicate an earlier row, comparing only the columns you name. ' +
      'The rows that survive move up to fill the gap. Column indices are counted from the first ' +
      'column of the range, not from column A.',
    mutates: true,
    parameters: {
      type: 'object',
      properties: {
        address: {
          type: 'string',
          description: 'The range to de-duplicate, including its header row, e.g. "A1:D40".',
        },
        columns: {
          type: 'array',
          description:
            'Which columns decide a duplicate, as indices within the range: 0 is the range\'s first ' +
            'column. Pass every column that must match, e.g. [0] to key on the first column only.',
          items: { type: 'number' },
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether the first row is a header row, which is never removed. Default true.',
        },
        sheet: SHEET,
      },
      required: ['address', 'columns'],
      additionalProperties: false,
    },
  },
  {
    name: 'web_search',
    description:
      'Search the public web for factual information you do not have: statistics, prices, specs, ' +
      'game data, dates, documentation, current events. Use this instead of asking the user to ' +
      'supply data you could look up yourself. Returns a short answer plus source snippets.',
    executesOn: 'server',
    mutates: false,
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'A focused search query. Be specific.' },
        maxResults: { type: 'integer', description: 'How many sources to return (1-10, default 6).' },
      },
      required: ['query'],
      additionalProperties: false,
    },
  },
];

/**
 * The tools to offer the model on a turn.
 *
 * Taskpane tools are limited to what the client advertised it can run (a stale
 * taskpane must never be sent a tool it doesn't know). Server tools are included
 * only when this server is actually configured to run them.
 */
export function toolsFor({
  capabilities,
  serverTools = [],
}: { capabilities?: string[] | null; serverTools?: string[] } = {}): ApiTool[] {
  const allow = capabilities && capabilities.length ? new Set(capabilities) : null;
  return TOOLS.filter((t) =>
    t.executesOn === 'server' ? serverTools.includes(t.name) : !allow || allow.has(t.name)
  ).map((t) => ({
    type: 'function' as const,
    function: { name: t.name, description: t.description, parameters: t.parameters },
  }));
}

export const SERVER_TOOL_NAMES: string[] = TOOLS.filter((t) => t.executesOn === 'server').map((t) => t.name);

export function getTool(name: string): ToolDefinition | null {
  return TOOLS.find((t) => t.name === name) || null;
}
