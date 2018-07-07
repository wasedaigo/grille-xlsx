"use strict";

var _ = require("lodash");
var xlsx = require("xlsx");

var EmptyColumnRegex = /__EMPTY.*/;

var GrilleXlsx = function(filename, metaTableName = 'meta') {
  this.filename = filename;
  this.metaTableName = metaTableName;
};

GrilleXlsx.prototype.toJson = function() {
  return GrilleXlsx.xlsxToJson(this.filename, this.metaTableName);
};

GrilleXlsx.getMetaData = function(workbook, metaTableName) {
  var sheet = workbook.Sheets[metaTableName];
  var json = xlsx.utils.sheet_to_json(sheet, { raw: true, blankrows: false });

  // Remove Empty cloumns from the meta data
  for (var i in json) {
    for (var j in json[i]) {
      if (EmptyColumnRegex.test(j)) {
        delete json[i][j];
      }
    }
  }
  return json;
};

GrilleXlsx.error = function(message, columnName, row, data) {
  return new Error(
    `${message} [column=${columnName} row=${row} data=${JSON.stringify(data)}]`
  );
};

GrilleXlsx.convert = function(columnName, type, row, data) {
  var result;
  switch (type) {
    case "ignore":
      result = "";

      break;

    case "integer":
      data = data || 0;
      result = Math.floor(data);

      break;

    case "float":
      data = data || 0.0;
      result = parseFloat(data);

      break;

    case "string":
      data = data || "";
      data = data + '';
      result = data;
      break;

    case "boolean":
      data = data || false;
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
      data = data || [];
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
      data = data || [];
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
      data = data || [];
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
      data = data || [];
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
      data = data || [];
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
      data = data || {};
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

GrilleXlsx.processWorksheet = function(workbook, metaTableName) {
  var meta = GrilleXlsx.getMetaData(workbook, metaTableName);

  var result = {};
  const keys = Object.keys(meta);
  for (var i = 0; i < keys.length; i++) {
    var def = meta[i];
    var sheetName = def.name;
    var path = def.collection;
    var sheet = workbook.Sheets[def.name];
    var json = xlsx.utils.sheet_to_json(sheet, { raw: true, blankrows: false });
    if (json.length === 0) {
      throw new Error("sheetName=" + sheetName + " does not exist or empty");
    }
    var header = Object.keys(json[0]).filter(function(k) {
      return !EmptyColumnRegex.test(k);
    });
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

GrilleXlsx.xlsxToJson = function(filename, meta) {
  var workbook = xlsx.readFile(filename);
  return GrilleXlsx.processWorksheet(workbook, meta);
};

module.exports = GrilleXlsx;
