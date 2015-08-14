# googlescript-markdowntablefive

### SYNOPSIS

```
  + MarkdownTableMaker();
  + toString() : string
  + getMarkdown() : string
  + getRange() : Range
  + setRange(Range) : MarkdownTableMaker
```

### DEMO

```javascript

function demo() {

  // -- SIMPLE -- //

  // builds object, selects entire sheet, crops data
  // builds markdown table
  var mtm = MarkdownTableMaker();

  // prints markdown table
  Logger.log('' + mtm);


  // -- A LITTLE MORE FUN -- //

  // selects user-specified range
  var range1 = SpreadsheetApp.getActive().getRange('B8:D19');

  // set input range, evaluates input range, crops data (if needed)
  // builds new table with bounded data
  mtm.setRange(range1);

  // prints new markdown table
  Logger.log('' + mtm);


}

```