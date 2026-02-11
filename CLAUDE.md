# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

mdbook-bom is an mdBook preprocessor that generates bill of materials (BOM), tools lists, and consumables lists from YAML front matter in assembly instruction chapters. It processes markdown files, extracts component information, inserts tables into the rendered book, and generates consolidated Excel output files.

## Build and Development Commands

```bash
# Build the project
cargo build

# Build with release optimizations
cargo build --release

# Install locally
cargo install --path .

# Run tests
cargo test

# Run with verbose output
cargo run -- supports html

# Check for compilation errors without building
cargo check

# Format code
cargo fmt

# Run clippy for linting
cargo clippy
```

## Configuration

The preprocessor is configured in `book.toml` of the mdBook project:

```toml
[preprocessor.bom]
output_path = "path/to/output.xlsx"  # Required - where to write the BOM Excel file
inventory_file = "~/path/to/inventory.xlsx"  # Required - Excel inventory file path
```

### Inventory Source

The preprocessor requires an Excel inventory file with the following sheets:
- Hardware
- Electronics
- Custom Parts
- Consumables
- Tools
- Assemblies

The `inventory_file` path supports home directory expansion with `~/`.

## Architecture

### Main Components

1. **BomPreprocessor**: Implements mdBook's `Preprocessor` trait
   - Entry point: `run()` method
   - Loads inventory data (Excel)
   - Processes all book chapters
   - Generates final BOM Excel output

2. **Inventory Loading**: Handles reading inventory data from Excel
   - `Inventory::load()` - Loads Excel file with home directory expansion
   - `load_*_from_excel()` methods - Read individual sheets (Hardware, Electronics, Custom Parts, Consumables, Tools, Assemblies)
   - Stores data in HashMaps keyed by part number/name

3. **Front Matter Processing**:
   - `extract_front_matter()` - Extracts YAML between `---` delimiters
   - `remove_front_matter()` - Strips YAML from content
   - Uses section-based structure only

4. **Table Generation**:
   - `insert_section_tables()` - Inserts collapsible tables after step headers
   - `generate_overview_tables()` - Creates chapter overview with all components
   - Generates separate tables for: Hardware, Electronics, Custom Parts, Consumables, Tools, Assemblies
   - Each table has collapsible `<details>` elements with unique IDs for JavaScript control

5. **BOM Accumulation**:
   - `accumulate_*()` functions - Aggregate components across all chapters
   - Combines quantities for parts and assemblies
   - Deduplicates consumables and tools
   - Merges tool settings from different steps
   - Respects `exclude_from_bom` flag — items with `exclude_from_bom: true` appear in chapter tables but are skipped during BOM accumulation

6. **Output Generation**:
   - `generate_bom_excel_file()` - Creates multi-sheet Excel workbook
   - Separate sheets for each component category (Hardware, Electronics, Custom Parts, Tools, Consumables, Assemblies)
   - Sorted by description/name

### Data Flow

```
book.toml config → load Excel inventory → process chapters →
  extract front matter → parse YAML → insert tables in markdown →
  accumulate all components (respecting exclude_from_bom) → generate Excel BOM
```

### Front Matter Structure
```yaml
---
sections:
  step_1:
    hardware:
      - name: "PART-001"
        quantity: 2
      - name: "SPECIAL-BOLT"
        quantity: 1
        exclude_from_bom: true   # Optional, defaults to false
    electronics:
      - name: "ELEC-001"
        quantity: 1
    custom_parts:
      - name: "CUSTOM-001"
        quantity: 1
    consumables:
      - name: "CONSUMABLE-001"
    tools:
      - name: "TOOL-001"
        setting: "5 Nm"  # Optional
    assemblies:
      - name: "Spool Module"
        quantity: 1
        exclude_from_bom: true   # Built in the book, not ordered
      - name: "Wire Harness"
        quantity: 2               # exclude_from_bom defaults to false → included in BOM
---
```

### `exclude_from_bom` Field

All component reference types (`PartReference`, `ConsumableReference`, `ToolReference`, `AssemblyReference`) support an optional `exclude_from_bom: bool` field (defaults to `false`). When set to `true`:
- The item still appears in chapter-level and overview tables
- The item is **skipped** during BOM accumulation (not included in the Excel output)
- Useful for sub-assemblies built within the book or items that should not be double-counted

In overview table deduplication, `exclude_from_bom` uses logical AND: only excluded if ALL references across sections exclude it.

### Step Header Matching

Step headers are matched using regex:
- Pattern: `(?i)^##+\s+Step\s+(\d+):?.*$`
- Matches: `## Step 1:` or `## Step 1` or `### Step 2:`
- Case-insensitive
- Section key format: `step_{number}` (e.g., `step_1` matches "Step 1")

### Inventory Data Structures

- `InventoryFastener`, `InventoryElectronic`, `InventoryCustomPart`: Have `Name` and optional `Description`
- `InventoryConsumable`: Has `Name` and optional `Description`
- `InventoryTool`: Has `Name` and optional `Brand`
- `InventoryAssembly`: Has `Name` and optional `Description`
- All use serde `#[serde(rename = "Name")]` to match Excel column headers

### BOM Data Structures

- `BomFastenerItem`, `BomElectronicItem`, `BomCustomPartItem`: Track `total_quantity` across all chapters
- `BomConsumableItem`: No quantity (treated as binary - needed or not)
- `BomToolItem`: Aggregates multiple `settings` from different chapters
- `BomAssemblyItem`: Tracks `total_quantity` across all chapters

## Key Implementation Details

- **Home directory expansion**: Excel paths starting with `~/` are expanded to full paths
- **Error handling**: Uses mdbook's `Error` type, wrapping underlying errors with context
- **Show All button**: JavaScript function `toggleAllTables()` controls `<details>` elements
- **Horizontal rules**: Inserted between steps (not before first step) for visual separation
- **Overview tables**: Inserted after the top-level `#` header before any steps

## Common Modifications

- **Adding new component categories**: Add new inventory struct, reference struct, BOM struct, accumulation function, table generation function, combine/dedup function, and sheet in Excel output
- **Changing table formatting**: Modify `generate_*_table()` functions
- **Adjusting step matching**: Update regex in `find_step_headers()`
- **Excel schema changes**: Update serde field mappings in inventory structs
