# JSONEncode

2022: Couldn't find my own functions so modded this and made changes to use a string buffer for a massive speed increase.

Functions for encoding strings, dates, arrays and dictionaries for JSON in Classic ASP. String encoding uses fast JavaScript regular expression search/replace.

## Usage

See test/index.asp for examples of usage.

### JSONEncodeString

Encode a string for embedding as a value in a JSON document

    t = JSONEncodeString(s)

### JSONEncodeDate

Format a date object into ISO-8601 so that JavaScript will parse it

    d = JSONEncodeDate(Now)

### JSONEncodeDict

Convert a dictionary object into a JSON object literal, with a parent element

    json = JSONEncodeDict("response", dict)

### JSONEncodeArray

Convert an array into a JSON array literal

    json = JSONEncodeArray(arr)
