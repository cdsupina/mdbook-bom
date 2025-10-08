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

The `inventory_file` path supports home directory expansion with `~/`.

## Architecture

### Main Components

1. **BomPreprocessor** (line 317-472): Implements mdBook's `Preprocessor` trait
   - Entry point: `run()` method
   - Loads inventory data (Excel or CSV)
   - Processes all book chapters
   - Generates final BOM Excel output

2. **Inventory Loading** (line 53-230): Handles reading inventory data from Excel
   - `Inventory::load()` - Loads Excel file with home directory expansion
   - `load_*_from_excel()` methods - Read individual sheets (Hardware, Electronics, Custom Parts, Consumables, Tools)
   - Stores data in HashMaps keyed by part number/name

3. **Front Matter Processing**:
   - `extract_front_matter()` - Extracts YAML between `---` delimiters
   - `remove_front_matter()` - Strips YAML from content
   - Uses section-based structure only

4. **Table Generation**:
   - `insert_section_tables()` (line 639-761) - Inserts collapsible tables after step headers
   - `generate_overview_tables()` (line 763-853) - Creates chapter overview with all components
   - Generates separate tables for: Hardware 🔩, Electronics 🔌, Custom Parts ⚙️, Consumables 🧪, Tools 🔧
   - Each table has collapsible `<details>` elements with unique IDs for JavaScript control

5. **BOM Accumulation**:
   - `accumulate_*()` functions (lines 1108-1246) - Aggregate components across all chapters
   - Combines quantities for parts
   - Deduplicates consumables and tools
   - Merges tool settings from different steps

6. **Output Generation**:
   - `generate_bom_excel_file()` (line 1261-1444) - Creates multi-sheet Excel workbook
   - Separate sheets for each component category
   - Sorted by description/name

### Data Flow

```
book.toml config → load Excel inventory → process chapters →
  extract front matter → parse YAML → insert tables in markdown →
  accumulate all components → generate Excel BOM
```

### Front Matter Structure
```yaml
---
sections:
  step_1:
    hardware:
      - name: "PART-001"
        quantity: 2
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
---
```

### Step Header Matching

Step headers are matched using regex (line 624):
- Pattern: `(?i)^##+\s+Step\s+(\d+):?.*$`
- Matches: `## Step 1:` or `## Step 1` or `### Step 2:`
- Case-insensitive
- Section key format: `step_{number}` (e.g., `step_1` matches "Step 1")

### Inventory Data Structures

- `InventoryFastener`, `InventoryElectronic`, `InventoryCustomPart`: Have `Name` and optional `Description`
- `InventoryConsumable`: Has `Name` and optional `Description`
- `InventoryTool`: Has `Name` and optional `Brand`
- All use serde `#[serde(rename = "Name")]` to match Excel column headers

### BOM Data Structures

- `BomFastenerItem`, `BomElectronicItem`, `BomCustomPartItem`: Track `total_quantity` across all chapters
- `BomConsumableItem`: No quantity (treated as binary - needed or not)
- `BomToolItem`: Aggregates multiple `settings` from different chapters

## Key Implementation Details

- **Home directory expansion**: Excel paths starting with `~/` are expanded to full paths (line 87-99)
- **Error handling**: Uses mdbook's `Error` type, wrapping underlying errors with context
- **Show All button**: JavaScript function `toggleAllTables()` controls `<details>` elements (line 926-947)
- **Horizontal rules**: Inserted between steps (not before first step) for visual separation (line 668-675)
- **Overview tables**: Inserted after the top-level `#` header before any steps (line 654-665)

## Common Modifications

- **Adding new component categories**: Add new inventory struct, accumulation function, table generation function, and sheet in Excel output
- **Changing table formatting**: Modify `generate_*_table()` functions (lines 953-1106)
- **Adjusting step matching**: Update regex in `find_step_headers()` (line 624)
- **Excel schema changes**: Update serde field mappings in inventory structs (lines 514-552)
