"use strict";

var assert = require("assert");

var GrilleXlsx = require("../lib/grille-xlsx.js");

describe("GrilleXlsx", function() {
  describe("check", function() {
    var json;
    before(function() {
      var grilleXlsx = new GrilleXlsx("test/test.xlsx");
      json = grilleXlsx.toJson();
    });
    it("testArray", function() {
      assert.equal(json.arrayCollection.length, 4);
      var sample = json.arrayCollection[3];
      assert.equal(sample.id, 3);
      assert.equal(sample.bool, false);
      assert.deepEqual(sample.floatArray, [1.0, 3.0, 4.0]);
      assert.deepEqual(sample.boolArray, [true, false, false]);
      assert.deepEqual(sample.stringArray, ["a", "b", "c"]);
    });
    it("testHash", function() {
      var keys = Object.keys(json.hashCollection);
      assert.deepEqual(keys, ["bronze_sword", "silver_sword", "gold_sword"]);
      var sample = json.hashCollection["silver_sword"];
      assert.equal(Object.keys(sample).length, 4);
      assert.equal(sample.atk, 1);
      assert.equal(sample.def, 1);
      assert.equal(sample.critical, 0.2);
    });
    it("testKeyValue", function() {
      var keys = Object.keys(json.keyvalueCollection);
      assert.deepEqual(keys, ["FLOAT", "INTEGER", "ARRAY_INTEGER", "STRING"]);

      assert.equal(json.keyvalueCollection["FLOAT"], 0.1);
      assert.equal(json.keyvalueCollection["INTEGER"], 10);
      assert.deepEqual(json.keyvalueCollection["ARRAY_INTEGER"], [1, 2, 3]);
      assert.equal(json.keyvalueCollection["STRING"], "test");
    });
    it("testArrayMerge", function() {
      assert.equal(json.arrayMerge.length, 5);
      console.log(JSON.stringify(json.arrayMerge));
      var sample = json.arrayMerge[4];
      assert.equal(sample.id, 5);
      assert.equal(sample.bool, true);
      assert.deepEqual(sample.floatArray, [1.0, 3.0, 4.0]);
      assert.deepEqual(sample.boolArray, [true, false, false]);
      assert.deepEqual(sample.stringArray, ["a", "b", "c"]);
    });
  });
});
