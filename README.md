# mdbook-bom

A preprocessor for [mdBook](https://rust-lang.github.io/mdBook/) that generates bill of materials (BOM), tools lists, and consumables lists from YAML front matter in assembly instruction chapters.

## Features

- **Section-specific metadata**: Define parts, tools, and consumables per assembly step
- **Automatic table generation**: Collapsible tables are inserted after step headers in the rendered book
- **BOM generation**: Creates consolidated Excel workbook with sheets for each component category
- **Inventory lookup**: Uses Excel inventory file with multiple sheets for component details
- **Flexible step matching**: Supports both `## Step 1:` and `## Step 1` header formats
- **Interactive UI**: Show All/Hide All buttons to toggle component tables visibility

## Installation

```bash
cargo install --git https://github.com/cdsupina/mdbook-bom.git
```

## Usage

### 1. Configure mdbook

Add the preprocessor to your `book.toml`:

```toml
[preprocessor.bom]
inventory_file = "~/path/to/inventory.xlsx"  # Required - path to Excel inventory
output_path = "output/BOM.xlsx"              # Required - where to write BOM output
```

### 2. Create inventory file

Create an Excel inventory file with the following sheets:

**Hardware sheet:**
| Name | Description |
|------|-------------|
| SCREW-M4x20 | M4x20mm socket head cap screw |

**Electronics sheet:**
| Name | Description |
|------|-------------|
| LED-RED-5MM | 5mm red LED |

**Custom Parts sheet:**
| Name | Description |
|------|-------------|
| BRACKET-001 | Custom mounting bracket |

**Consumables sheet:**
| Name | Description |
|------|-------------|
| THREADLOCK-242 | Loctite 242 threadlocker |

**Tools sheet:**
| Name | Brand |
|------|-------|
| ALLEN-4MM | Wiha |

### 3. Add front matter to chapters

Add YAML front matter to your markdown chapters:

```yaml
---
sections:
  step_1:
    consumables:
      - name: "THREADLOCK-242"
    tools:
      - name: "SAFETY-GLASSES"
  step_2:
    hardware:
      - name: "SCREW-M4x20"
        quantity: 4
    electronics:
      - name: "LED-RED-5MM"
        quantity: 2
    custom_parts:
      - name: "BRACKET-001"
        quantity: 1
    tools:
      - name: "ALLEN-4MM"
      - name: "TORQUE-WRENCH"
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
- Insert collapsible requirement tables after each step header
- Insert an overview table at the top of each chapter with all components needed
- Generate Excel workbook at the specified `output_path` with consolidated BOM

## Output Files

The preprocessor generates an Excel workbook with separate sheets:

- **Hardware**: All hardware/fasteners with quantities
- **Electronics**: All electronic components with quantities
- **Custom Parts**: All custom parts with quantities
- **Tools**: All required tools with brands (settings not included in BOM)
- **Consumables**: All consumables needed

## Front Matter Structure

```yaml
sections:
  step_1:
    hardware:
      - name: "SCREW-M4x20"
        quantity: 2
    electronics:
      - name: "LED-RED-5MM"
        quantity: 1
    custom_parts:
      - name: "BRACKET-001"
        quantity: 1
    consumables:
      - name: "THREADLOCK-242"
    tools:
      - name: "ALLEN-4MM"
        setting: "5 Nm"  # Optional setting
```

All fields (hardware, electronics, custom_parts, consumables, tools) are optional for each step.

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