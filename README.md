# mdbook-bom

A preprocessor for [mdBook](https://rust-lang.github.io/mdBook/) that generates bill of materials (BOM), tools lists, and consumables lists from YAML front matter in assembly instruction chapters.

## Features

- **Section-specific metadata**: Define parts, tools, and consumables per assembly step
- **Automatic table generation**: Tables are inserted after step headers in the rendered book
- **BOM generation**: Creates consolidated CSV files for parts, tools, and consumables
- **Inventory lookup**: Uses separate CSV inventory files for part details
- **Flexible step matching**: Supports both `## Step 1:` and `## Step 1` header formats
- **Backwards compatibility**: Supports legacy chapter-level front matter

## Installation

```bash
cargo install --git https://github.com/cdsupina/mdbook-bom.git
```

## Usage

### 1. Configure mdbook

Add the preprocessor to your `book.toml`:

```toml
[preprocessor.bom]
```

### 2. Create inventory files

Create CSV inventory files in your project:

**inventory/parts.csv:**
```csv
part_number,description,supplier,unit_cost
SCREW-001,M4x20mm socket head cap screw,McMaster-Carr,0.15
```

**inventory/consumables.csv:**
```csv
part_number,description,supplier,unit_cost
THREADLOCK-001,Loctite 242 threadlocker,McMaster-Carr,8.50
```

**inventory/tools.csv:**
```csv
part_number,description,supplier
ALLEN-4MM,4mm hex allen key,McMaster-Carr
```

### 3. Add front matter to chapters

Use section-based YAML front matter in your markdown chapters:

```yaml
---
sections:
  step_1:
    consumables:
      - part_number: "CLEANER-001"
    tools:
      - part_number: "SAFETY-GLASSES"
  step_2:
    parts:
      - part_number: "SCREW-001"
        quantity: 4
    tools:
      - part_number: "ALLEN-4MM"
      - part_number: "TORQUE-WRENCH"
        setting: "5 Nm"
---

# Chapter 1: Assembly

## Step 1: Preparation
Instructions here...

## Step 2: Main Assembly
More instructions...
```

### 4. Build your book

```bash
mdbook build
```

The preprocessor will:
- Insert requirement tables after each step header
- Generate `output/BOM.csv` with consolidated parts list
- Generate `output/tools.csv` with required tools
- Generate `output/consumables.csv` with consumables list

## Output Files

The preprocessor generates three CSV files in the `output/` directory:

- **BOM.csv**: Parts with quantities and costs
- **tools.csv**: Tools required (without settings column)  
- **consumables.csv**: Consumables with costs

## Front Matter Structure

### Section-based (recommended)
```yaml
sections:
  step_1:
    parts:
      - part_number: "PART-001"
        quantity: 2
    consumables:
      - part_number: "CONSUMABLE-001"
    tools:
      - part_number: "TOOL-001"
        setting: "5 Nm"  # Optional setting
```

### Legacy chapter-level (backwards compatible)
```yaml
parts:
  - part_number: "PART-001"
    quantity: 2
tools:
  - part_number: "TOOL-001"
```

## Step Header Matching

The preprocessor matches section keys to markdown headers:
- `step_1` matches `## Step 1:` or `## Step 1`
- `step_2` matches `## Step 2:` or `## Step 2`
- Case-insensitive matching

## Requirements

- Rust 1.70+
- mdBook 0.4+

## License

MIT License