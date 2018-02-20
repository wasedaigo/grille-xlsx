# grille-xlsx-to-json
## Usage
```
  var grilleXlsx = new GrilleXlsx("test/test.xlsx");
  var json = grilleXlsx.toJson();
```

## Data Types

Grille supports the following list of data types:

Name            | Examples
----------------|----------------------
integer         | 1, -2, 99999
json            | [1, 2, 3], {"a": "b"}
string          | Banana
boolean         | TRUE/FALSE
float           | 1.2, 99.9, 2
array           | [1, true, "blah"]
array.integer   | [1, 2, 3]
array.string    | ["first", "second"]
array.boolean   | [true, false]
array.float     | [1, 1.1, 1.2]

I recommend using data validation on the second row of a worksheet to enforce these (see example spreadsheet).

### Example Meta Worksheet

The `meta` worksheet tells Grille how to parse your content.
It is loaded prior to all other sheets being loaded.

The `id` column correlates to the worksheet (tab) name to be loaded (if it's not listed it's not loaded).

The `collection` column tells Grille which top-level attribute the data for that worksheet should be stored at. Note that you can use `.` for specifying deeper nested objects.

The `format` column tells Grille which method to use when converting the raw worksheet into a native object.

As a convention, all worksheets specify data types as the second row. I suggest using Data Validation (like in the example worksheet).

id                  | collection    | format
--------------------|---------------|---------
string              | string        | string
people              | people        | hash
keyvalue\_string    | keyvalue      | keyvalue
keyvalue\_integer   | keyvalue      | keyvalue
level\_1            | levels.0      | array
level\_2            | levels.1      | array
level\_secret       | levels.secret | array


### Example Array Worksheet

This will likely be the most common format you use.
Data is loaded into an array where each reacord is corresponding to each row of the speadsheet.


id      | name              | likesgum  | gender
--------|-------------------|-----------|------
integer | string            | boolean   | string
1       | Rupert Styx       | FALSE     | m
2       | Morticia Addams   | TRUE      | f

#### Hash Output

```json
{
  "people": {
    "1": {
      "gender": "m",
      "id": 1,
      "likesgum": false,
      "name": "Rupert Styx"
    },
    "2": {
      "gender": "f",
      "id": 2,
      "likesgum": true,
      "name": "Morticia Addams"
    }
  }
}
```

### Example Hash Worksheet

Data is loaded into an object where each key is the value in the `id` column.
The `id` column should be a number or a string and each row should have a unique value.


id      | name              | likesgum  | gender
--------|-------------------|-----------|------
integer | string            | boolean   | string
1       | Rupert Styx       | FALSE     | m
2       | Morticia Addams   | TRUE      | f

#### Hash Output

```json
{
  "people": {
    "1": {
      "gender": "m",
      "id": 1,
      "likesgum": false,
      "name": "Rupert Styx"
    },
    "2": {
      "gender": "f",
      "id": 2,
      "likesgum": true,
      "name": "Morticia Addams"
    }
  }
}
```

### Example KeyValue Worksheet

KeyValue worksheets provide a simple collection for looking up data.

Since each worksheet can only contain a single data type, I recommend using multiple sheets for different types and merging them together.
Simply set the resulting `meta` collections for multiple sheets to be the same (see above) and they will be merged together as expected.

id          | value
------------|-----------------
string      | string
title       | Simple CMS Demo
author      | Thomas Hunter II

#### KeyValue Output

```json
{
  "keyvalue": {
    "author": "Thomas Hunter II",
    "title": "Simple CMS Demo"
  }
}
```