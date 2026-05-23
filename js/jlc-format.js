(function (global) {
  "use strict";

  const JLC_FORMAT = {
    bom: {
      id: "bom",
      label: "BOM (Bill of Materials)",
      description: "Comment, Designator, Footprint",
      columns: [
        {
          key: "comment",
          header: "Comment",
          required: true,
          aliases: [
            "comment",
            "value",
            "description",
            "comp value",
            "component value",
            "part value",
            "name",
            "val",
          ],
        },
        {
          key: "designator",
          header: "Designator",
          required: true,
          aliases: [
            "designator",
            "designators",
            "ref",
            "refdes",
            "reference",
            "references",
            "ref des",
          ],
        },
        {
          key: "footprint",
          header: "Footprint",
          required: true,
          aliases: [
            "footprint",
            "package",
            "pattern",
            "fp",
            "footprint name",
            "land pattern",
          ],
        },
      ],
      fileSuffix: "_JLC_BOM",
    },
    pnp: {
      id: "pnp",
      label: "Pick & Place",
      description: "Designator, Mid X, Mid Y, Layer, Rotation",
      columns: [
        {
          key: "designator",
          header: "Designator",
          required: true,
          aliases: [
            "designator",
            "ref",
            "refdes",
            "reference",
            "ref des",
          ],
        },
        {
          key: "midX",
          header: "Mid X",
          required: true,
          aliases: [
            "mid x",
            "mid x mm",
            "x",
            "x mm",
            "x (mm)",
            "posx",
            "pos x",
            "center x",
          ],
        },
        {
          key: "midY",
          header: "Mid Y",
          required: true,
          aliases: [
            "mid y",
            "mid y mm",
            "y",
            "y mm",
            "y (mm)",
            "posy",
            "pos y",
            "center y",
          ],
        },
        {
          key: "layer",
          header: "Layer",
          required: true,
          aliases: ["layer", "side", "board side", "top bottom"],
        },
        {
          key: "rotation",
          header: "Rotation",
          required: true,
          aliases: [
            "rotation",
            "rot",
            "rotate",
            "angle",
            "orientation",
          ],
        },
      ],
      fileSuffix: "_JLC_PickPlace",
    },
  };

  function normalizeHeader(h) {
    return String(h ?? "")
      .toLowerCase()
      .replace(/[（）()]/g, " ")
      .replace(/[^a-z0-9#]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function escapeCSVField(value) {
    const s = String(value ?? "");
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function rowsToCSV(rows) {
    return rows.map((r) => r.map(escapeCSVField).join(",")).join("\r\n");
  }

  function findColumnIndex(headers, aliases) {
    const normalized = headers.map(normalizeHeader);
    for (const alias of aliases) {
      const exact = normalized.findIndex((h) => h === alias);
      if (exact >= 0) return exact;
    }
    for (const alias of aliases) {
      const partial = normalized.findIndex(
        (h) => h.includes(alias) || alias.includes(h)
      );
      if (partial >= 0) return partial;
    }
    return -1;
  }

  function autoMap(headers, formatId) {
    const format = JLC_FORMAT[formatId];
    if (!format) return {};
    const mapping = {};
    format.columns.forEach((col) => {
      const idx = findColumnIndex(headers, col.aliases);
      mapping[col.key] = idx >= 0 ? idx : -1;
    });
    return mapping;
  }

  function formatCoord(value) {
    const s = String(value ?? "").trim();
    if (!s) return "";
    if (/mm\s*$/i.test(s)) return s.replace(/\s*mm\s*$/i, "mm");
    const n = parseFloat(s.replace(/[^0-9.\-eE+]/g, ""));
    if (Number.isNaN(n)) return s;
    return n + "mm";
  }

  function normalizeLayer(value) {
    const v = String(value ?? "")
      .trim()
      .toLowerCase();
    if (!v) return "Top";
    if (
      v.includes("bottom") ||
      v === "b" ||
      v === "bot" ||
      v === "2" ||
      v === "back"
    ) {
      return "Bottom";
    }
    if (v.includes("top") || v === "t" || v === "1" || v === "front") {
      return "Top";
    }
    return v.charAt(0).toUpperCase() + v.slice(1);
  }

  function normalizeRotation(value) {
    const s = String(value ?? "").trim();
    if (!s) return "";
    const n = parseFloat(s.replace(/[^0-9.\-]/g, ""));
    if (Number.isNaN(n)) return s;
    return String(Math.round(n * 1000) / 1000);
  }

  function getCell(row, colIndex) {
    if (colIndex < 0 || colIndex >= row.length) return "";
    return String(row[colIndex] ?? "").trim();
  }

  function transformValue(formatId, key, raw) {
    if (formatId === "pnp") {
      if (key === "midX" || key === "midY") return formatCoord(raw);
      if (key === "layer") return normalizeLayer(raw);
      if (key === "rotation") return normalizeRotation(raw);
    }
    return raw;
  }

  function convert(headers, dataRows, formatId, mapping) {
    const format = JLC_FORMAT[formatId];
    if (!format) throw new Error("Unknown JLC format");

    const out = [format.columns.map((c) => c.header)];

    dataRows.forEach((row) => {
      const hasData = format.columns.some(
        (col) => getCell(row, mapping[col.key]) !== ""
      );
      if (!hasData) return;

      const outRow = format.columns.map((col) => {
        const raw = getCell(row, mapping[col.key]);
        return transformValue(formatId, col.key, raw);
      });
      out.push(outRow);
    });

    return out;
  }

  function validateMapping(formatId, mapping) {
    const format = JLC_FORMAT[formatId];
    const missing = format.columns
      .filter((c) => c.required && mapping[c.key] < 0)
      .map((c) => c.header);
    return missing;
  }

  global.JLCFormat = {
    FORMATS: JLC_FORMAT,
    autoMap,
    convert,
    rowsToCSV,
    validateMapping,
    normalizeHeader,
  };
})(typeof window !== "undefined" ? window : globalThis);
