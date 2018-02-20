"use strict";

var _ = require("lodash");
var xlsx = require("xlsx");

var EmptyColumnRegex = /__EMPTY.*/;

var GrilleXlsx = function(filename) {
  this.filename = filename;
};

GrilleXlsx.prototype.toJson = function() {
  return GrilleXlsx.xlsxToJson(this.filename);
};

GrilleXlsx.getMetaData = function(workbook) {
  var sheet = workbook.Sheets.meta;
  return xlsx.utils.sheet_to_json(sheet, { raw: true });
};

GrilleXlsx.error = function(message, columnName, row, data) {
  return new Error(
    `${message} [column=${columnName} row=${row} data=${JSON.stringify(data)}]`
  );
};

GrilleXlsx.convert = function(columnName, type, row, data) {
  if (data === undefined) {
    throw GrilleXlsx.error("Undefined data", columnName, row, data);
  }

  var result;
  switch (type) {
    case "ignore":
      result = "";

      break;

    case "integer":
      result = Math.floor(data);

      break;

    case "float":
      result = parseFloat(data);

      break;

    case "string":
      result = data;

      break;

    case "boolean":
      if (typeof data === "boolean") {
        result = data;
      } else {
        if (data !== "TRUE" && data !== "FALSE") {
          throw GrilleXlsx.error("Not a boolean", columnName, row, data);
        }

        result = data === "TRUE";
      }
      break;

    case "array":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }

        if (!Array.isArray(result)) {
          throw GrilleXlsx.error(
            "Data is not of type array",
            columnName,
            row,
            data
          );
        }
      } else {
        result = [];
      }

      break;

    case "array.integer":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }

        if (!Array.isArray(result)) {
          throw GrilleXlsx.error(
            "Data is not of type array",
            columnName,
            row,
            data
          );
        }

        result.forEach(function(value) {
          if (typeof value !== "number" || value % 1 !== 0) {
            throw new Error("Not an array of integers");
          }
        });
      } else {
        result = [];
      }

      break;

    case "array.string":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }

        if (!Array.isArray(result)) {
          throw GrilleXlsx.error(
            "Data is not of type array",
            columnName,
            row,
            data
          );
        }

        result.forEach(function(value) {
          if (typeof value !== "string") {
            throw GrilleXlsx.error(
              "Not an array of strings",
              columnName,
              row,
              data
            );
          }
        });
      } else {
        result = [];
      }

      break;

    case "array.boolean":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }

        if (!Array.isArray(result)) {
          throw GrilleXlsx.error(
            "Data is not of type array",
            columnName,
            row,
            data
          );
        }

        result.forEach(function(value) {
          if (typeof value !== "boolean") {
            throw GrilleXlsx.error(
              "Not an array of booleans",
              columnName,
              row,
              data
            );
          }
        });
      } else {
        result = [];
      }

      break;

    case "array.float":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }

        if (!Array.isArray(result)) {
          throw GrilleXlsx.error(
            "Data is not of type array",
            columnName,
            row,
            data
          );
        }

        result.forEach(function(value) {
          if (typeof value !== "number") {
            throw GrilleXlsx.error(
              "Not an array of floats",
              columnName,
              row,
              data
            );
          }
        });
      } else {
        result = [];
      }

      break;

    case "json":
      if (data) {
        try {
          result = JSON.parse(data);
        } catch (e) {
          throw GrilleXlsx.error("Unable to parse JSON", columnName, row, data);
        }
      } else {
        result = null;
      }

      break;

    default:
      throw GrilleXlsx.error(
        `Unable to parse data type [type] ${type}`,
        columnName,
        row,
        data
      );
  }

  return result;
};

GrilleXlsx.convertJson = function(sheetName, header, typeMap, json) {
  var result = [];
  for (var i = 0; i < json.length; i++) {
    var row = {};
    for (var j = 0; j < header.length; j++) {
      var k = header[j];
      row[k] = GrilleXlsx.convert(
        sheetName + ":" + k,
        typeMap[k],
        i,
        json[i][k]
      );
    }
    result.push(row);
  }
  return result;
};

GrilleXlsx.processArray = function(result, sheetName, header, json, path) {
  var typeMap = json.shift();
  var conent = GrilleXlsx.convertJson(sheetName, header, typeMap, json);
  GrilleXlsx.dotSet(result, conent, path);
};

GrilleXlsx.processHash = function(result, sheetName, header, json, path) {
  var typeMap = json.shift();
  var conent = GrilleXlsx.convertJson(sheetName, header, typeMap, json);
  GrilleXlsx.dotSet(result, _.keyBy(conent, "id"), path);
};

GrilleXlsx.processKeyValue = function(result, sheetName, header, json, path) {
  GrilleXlsx.dotMerge(
    result,
    GrilleXlsx.extractKeyValue(sheetName, json),
    path
  );
};

GrilleXlsx.processWorksheet = function(workbook) {
  var meta = GrilleXlsx.getMetaData(workbook);

  var result = {};
  for (var key in meta) {
    var def = meta[key];
    var sheet = workbook.Sheets[def.name];
    var json = xlsx.utils.sheet_to_json(sheet, { raw: true });
    if (json.length === 0) return result;
    var header = Object.keys(json[0]).filter(function(k) {
      return !EmptyColumnRegex.test(k);
    });
    var sheetName = def.name;
    var path = def.collection;
    switch (def.format) {
      case "array":
        GrilleXlsx.processArray(result, sheetName, header, json, path);
        break;
      case "hash":
        GrilleXlsx.processHash(result, sheetName, header, json, path);
        break;
      case "keyvalue":
        GrilleXlsx.processKeyValue(result, sheetName, header, json, path);
        break;
    }
  }

  return result;
};

GrilleXlsx.extractKeyValue = function(sheetName, keyvalues) {
  var result = {};
  for (var i = 0; i < keyvalues.length; i++) {
    var keyvalue = keyvalues[i];
    result[keyvalue.key] = GrilleXlsx.convert(
      sheetName,
      keyvalue.type,
      i,
      keyvalue.value
    );
  }

  return result;
};

/**
 * Sets data in a deep object using dot notation
 */
GrilleXlsx.dotSet = function(destination, source, path) {
  var nodes = path.split("."); // e.g. x, y, z

  var pointer = destination;

  for (var i = 0; i < nodes.length - 1; i++) {
    if (!pointer[nodes[i]]) {
      pointer[nodes[i]] = {};
    }

    pointer = pointer[nodes[i]];
  }

  pointer[nodes[nodes.length - 1]] = source;

  return destination;
};

/**
 * Merges data in a deep object using dot notation
 */
GrilleXlsx.dotMerge = function(destination, source, path) {
  var nodes = path.split("."); // e.g. x, y, z

  var pointer = destination;

  for (var i = 0; i < nodes.length; i++) {
    if (!pointer[nodes[i]]) {
      pointer[nodes[i]] = {};
    }

    pointer = pointer[nodes[i]];
  }

  _.extend(pointer, source);

  return destination;
};

GrilleXlsx.xlsxToJson = function(filename) {
  var workbook = xlsx.readFile(filename);
  return GrilleXlsx.processWorksheet(workbook);
};

module.exports = GrilleXlsx;
