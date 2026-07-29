# Excel Sprite Catalog Engine

Generate sprite catalogs from Excel worksheet data.

Excel Sprite Catalog Engine is a VBA-based engine that converts
worksheet data into sprite codes and generates a catalog for bitmap rendering.

This engine works as a bridge between Excel worksheet data and
Excel Sprite Engine.

## Overview

Excel Sprite Catalog Engine converts worksheet data into
sprite catalogs for bitmap resource management.

This engine extracts unique characters from a worksheet area
and generates a catalog that can be used by sprite-based applications.

## Features

- Convert numbers and alphabet characters into sprite codes
- Generate unique sprite catalogs automatically
- Prepare bitmap resources for Excel Sprite Engine
- Work directly with worksheet cells

## Concept

Instead of displaying characters using standard Excel fonts,
this engine allows worksheet data to be transformed into
bitmap-based visual elements.

This makes it possible to create custom displays,
game interfaces, scoreboards, and other graphical applications
inside Excel.

## Workflow
```text
Worksheet Data
|
v
Replace Characters
|
v
Generate Sprite Catalog
```

## Sample

A sample workbook is included in the `sample` folder.

The sample demonstrates how to generate a sprite catalog
from worksheet data.

Included areas:

```
SCREEN AREA
    |
    v
Replace worksheet characters
    |
    v
CATALOG AREA
```

The sample workbook contains:

- A worksheet area for replacement processing
- A catalog area for generated sprite codes
- A test routine to execute the catalog generation process

Open the sample workbook and run `TEST_REPL`
to check the Replace Engine operation.

## Requirements

- Microsoft Excel with VBA support
- Windows environment

## License

This project is licensed under the MIT License.


