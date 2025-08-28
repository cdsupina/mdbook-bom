use calamine::{open_workbook, Reader, RangeDeserializerBuilder, Xlsx};
use clap::{Arg, ArgMatches, Command};
use mdbook::book::{Book, BookItem};
use mdbook::errors::Error;
use mdbook::preprocess::{CmdPreprocessor, Preprocessor, PreprocessorContext};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io;
use std::path::Path;

pub fn make_app() -> Command {
    Command::new("mdbook-bom")
        .about("A mdbook preprocessor to extract BOM from YAML front matter")
        .subcommand(
            Command::new("supports")
                .arg(Arg::new("renderer").required(true))
                .about("Check whether a renderer is supported by this preprocessor"),
        )
}

fn main() {
    let matches = make_app().get_matches();

    if let Some(sub_args) = matches.subcommand_matches("supports") {
        handle_supports(sub_args);
    } else if let Err(e) = handle_preprocessing() {
        eprintln!("{}", e);
        std::process::exit(1);
    }
}

fn handle_supports(sub_args: &ArgMatches) -> ! {
    let renderer = sub_args
        .get_one::<String>("renderer")
        .expect("Required argument");
    let supported = renderer != "not-supported";

    if supported {
        std::process::exit(0);
    } else {
        std::process::exit(1);
    }
}

fn handle_preprocessing() -> Result<(), Error> {
    let (ctx, book) = CmdPreprocessor::parse_input(io::stdin())?;
    let processed_book = BomPreprocessor.run(&ctx, book)?;
    serde_json::to_writer(io::stdout(), &processed_book)?;
    Ok(())
}

struct Inventory {
    fasteners: HashMap<String, InventoryFastener>,
    electronics: HashMap<String, InventoryElectronic>,
    custom_parts: HashMap<String, InventoryCustomPart>,
    consumables: HashMap<String, InventoryConsumable>,
    tools: HashMap<String, InventoryTool>,
}

impl Inventory {
    fn load(excel_path: Option<&str>) -> Result<Self, Error> {
        if let Some(path) = excel_path {
            Self::load_from_excel(path)
        } else {
            Self::load_from_csv()
        }
    }

    fn load_from_csv() -> Result<Self, Error> {
        let fasteners = Self::load_parts_as_fasteners()?;
        let electronics = HashMap::new(); // No electronics in legacy CSV
        let custom_parts = HashMap::new(); // No custom_parts in legacy CSV
        let consumables = Self::load_consumables()?;
        let tools = Self::load_tools()?;

        Ok(Inventory {
            fasteners,
            electronics,
            custom_parts,
            consumables,
            tools,
        })
    }

    fn load_from_excel(excel_path: &str) -> Result<Self, Error> {
        // Expand home directory if needed
        let expanded_path = if excel_path.starts_with("~/") {
            if let Some(home) = std::env::var_os("HOME") {
                let home_path = std::path::Path::new(&home);
                home_path.join(&excel_path[2..]).to_string_lossy().to_string()
            } else {
                return Err(Error::msg("Cannot expand ~ - HOME environment variable not set"));
            }
        } else {
            excel_path.to_string()
        };

        // Check if file exists first
        if !std::path::Path::new(&expanded_path).exists() {
            return Err(Error::msg(format!("Excel file not found: {}", expanded_path)));
        }

        let fasteners = Self::load_fasteners_from_excel(&expanded_path)?;
        let electronics = Self::load_electronics_from_excel(&expanded_path)?;
        let custom_parts = Self::load_custom_parts_from_excel(&expanded_path)?;
        let consumables = Self::load_consumables_from_excel(&expanded_path)?;
        let tools = Self::load_tools_from_excel(&expanded_path)?;

        Ok(Inventory {
            fasteners,
            electronics,
            custom_parts,
            consumables,
            tools,
        })
    }

    fn load_parts_as_fasteners() -> Result<HashMap<String, InventoryFastener>, Error> {
        let path = Path::new("inventory/parts.csv");
        if !path.exists() {
            return Err(Error::msg("inventory/parts.csv not found"));
        }

        let mut reader = csv::Reader::from_path(path)
            .map_err(|e| Error::msg(format!("Failed to read parts.csv: {}", e)))?;

        let mut fasteners = HashMap::new();
        for result in reader.deserialize() {
            let fastener: InventoryFastener =
                result.map_err(|e| Error::msg(format!("Failed to parse fastener: {}", e)))?;
            fasteners.insert(fastener.part_number.clone(), fastener);
        }

        Ok(fasteners)
    }

    fn load_consumables() -> Result<HashMap<String, InventoryConsumable>, Error> {
        let path = Path::new("inventory/consumables.csv");
        if !path.exists() {
            return Err(Error::msg("inventory/consumables.csv not found"));
        }

        let mut reader = csv::Reader::from_path(path)
            .map_err(|e| Error::msg(format!("Failed to read consumables.csv: {}", e)))?;

        let mut consumables = HashMap::new();
        for result in reader.deserialize() {
            let consumable: InventoryConsumable =
                result.map_err(|e| Error::msg(format!("Failed to parse consumable: {}", e)))?;
            consumables.insert(consumable.part_number.clone(), consumable);
        }

        Ok(consumables)
    }

    fn load_tools() -> Result<HashMap<String, InventoryTool>, Error> {
        let path = Path::new("inventory/tools.csv");
        if !path.exists() {
            return Err(Error::msg("inventory/tools.csv not found"));
        }

        let mut reader = csv::Reader::from_path(path)
            .map_err(|e| Error::msg(format!("Failed to read tools.csv: {}", e)))?;

        let mut tools = HashMap::new();
        for result in reader.deserialize() {
            let tool: InventoryTool =
                result.map_err(|e| Error::msg(format!("Failed to parse tool: {}", e)))?;
            tools.insert(tool.name.clone(), tool);
        }

        Ok(tools)
    }

    fn load_fasteners_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryFastener>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook.worksheet_range("hardware")
            .map_err(|e| Error::msg(format!("Failed to read 'hardware' sheet: {}", e)))?;

        let mut hardware = HashMap::new();
        let mut iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for hardware: {}", e)))?;

        for result in iter {
            let hardware_item: InventoryFastener = result
                .map_err(|e| Error::msg(format!("Failed to parse hardware row: {}", e)))?;
            hardware.insert(hardware_item.part_number.clone(), hardware_item);
        }

        Ok(hardware)
    }

    fn load_electronics_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryElectronic>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook.worksheet_range("electronics")
            .map_err(|e| Error::msg(format!("Failed to read 'electronics' sheet: {}", e)))?;


        let mut electronics = HashMap::new();
        let mut iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for electronics: {}", e)))?;

        for result in iter {
            let electronic: InventoryElectronic = result
                .map_err(|e| Error::msg(format!("Failed to parse electronic row: {}", e)))?;
            electronics.insert(electronic.part_number.clone(), electronic);
        }

        Ok(electronics)
    }

    fn load_custom_parts_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryCustomPart>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook.worksheet_range("custom_parts")
            .map_err(|e| Error::msg(format!("Failed to read 'custom_parts' sheet: {}", e)))?;

        let mut custom_parts = HashMap::new();
        let mut iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for custom_parts: {}", e)))?;

        for result in iter {
            let custom_part: InventoryCustomPart = result
                .map_err(|e| Error::msg(format!("Failed to parse custom_part row: {}", e)))?;
            custom_parts.insert(custom_part.part_number.clone(), custom_part);
        }

        Ok(custom_parts)
    }

    fn load_consumables_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryConsumable>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook.worksheet_range("consumables")
            .map_err(|e| Error::msg(format!("Failed to read 'consumables' sheet: {}", e)))?;

        let mut consumables = HashMap::new();
        let mut iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for consumables: {}", e)))?;

        for result in iter {
            let consumable: InventoryConsumable = result
                .map_err(|e| Error::msg(format!("Failed to parse consumable row: {}", e)))?;
            consumables.insert(consumable.part_number.clone(), consumable);
        }

        Ok(consumables)
    }

    fn load_tools_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryTool>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook.worksheet_range("tools")
            .map_err(|e| Error::msg(format!("Failed to read 'tools' sheet: {}", e)))?;

        let mut tools = HashMap::new();
        let mut iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for tools: {}", e)))?;

        for result in iter {
            let tool: InventoryTool = result
                .map_err(|e| Error::msg(format!("Failed to parse tool row: {}", e)))?;
            tools.insert(tool.name.clone(), tool);
        }

        Ok(tools)
    }
}

pub struct BomPreprocessor;

impl Preprocessor for BomPreprocessor {
    fn name(&self) -> &str {
        "bom"
    }

    fn run(&self, ctx: &PreprocessorContext, mut book: Book) -> Result<Book, Error> {
        // Check for Excel inventory file configuration
        let excel_path = if let Some(bom_cfg) = ctx.config.get_preprocessor(self.name()) {
            if let Some(inventory_file) = bom_cfg.get("inventory_file") {
                inventory_file.as_str()
            } else {
                None
            }
        } else {
            None
        };

        // Load inventory data
        let inventory = Inventory::load(excel_path)?;

        let mut all_fasteners: HashMap<String, BomFastenerItem> = HashMap::new();
        let mut all_electronics: HashMap<String, BomElectronicItem> = HashMap::new();
        let mut all_custom_parts: HashMap<String, BomCustomPartItem> = HashMap::new();
        let mut all_consumables: HashMap<String, BomConsumableItem> = HashMap::new();
        let mut all_tools: HashMap<String, BomToolItem> = HashMap::new();

        book.for_each_mut(|item: &mut BookItem| {
            if let BookItem::Chapter(ch) = item {
                if let Some(front_matter) = extract_front_matter(&ch.content) {
                    // Remove front matter from content
                    let content_without_fm = remove_front_matter(&ch.content);

                    // Parse YAML
                    if let Ok(metadata) = serde_yaml::from_str::<ChapterMetadata>(&front_matter) {
                        // Handle new section-based structure
                        if let Some(sections) = &metadata.sections {
                            // Insert tables after step headers
                            ch.content =
                                insert_section_tables(&content_without_fm, sections, &inventory);

                            // Accumulate all items from all sections for BOM
                            for section_metadata in sections.values() {
                                // Check both hardware and fasteners for backward compatibility
                                let hardware = section_metadata.hardware.as_deref().unwrap_or_default();
                                let legacy_fasteners = section_metadata.fasteners.as_deref().unwrap_or_default();
                                let electronics = section_metadata.electronics.as_deref().unwrap_or_default();
                                let custom_parts = section_metadata.custom_parts.as_deref().unwrap_or_default();
                                let consumables = section_metadata.consumables.as_deref().unwrap_or_default();
                                let tools = section_metadata.tools.as_deref().unwrap_or_default();
                                
                                // Legacy support: if parts exist, treat as fasteners for backward compatibility
                                let legacy_parts = section_metadata.parts.as_deref().unwrap_or_default();

                                accumulate_fasteners(hardware, &inventory, &mut all_fasteners);
                                accumulate_fasteners(legacy_fasteners, &inventory, &mut all_fasteners);
                                accumulate_fasteners(legacy_parts, &inventory, &mut all_fasteners); // Legacy support
                                accumulate_electronics(electronics, &inventory, &mut all_electronics);
                                accumulate_custom_parts(custom_parts, &inventory, &mut all_custom_parts);
                                accumulate_consumables(consumables, &inventory, &mut all_consumables);
                                accumulate_tools(tools, &inventory, &mut all_tools);
                            }
                        } else {
                            // Handle legacy flat structure (backwards compatibility)
                            ch.content = content_without_fm;

                            let parts = metadata.parts.as_deref().unwrap_or_default();
                            let consumables = metadata.consumables.as_deref().unwrap_or_default();
                            let tools = metadata.tools.as_deref().unwrap_or_default();

                            // Generate tables for this chapter (legacy behavior)
                            let parts_table = generate_fasteners_table(parts, &inventory, "legacy");
                            let consumables_table =
                                generate_consumables_table(consumables, &inventory, "legacy");
                            let tools_table = generate_tools_table(tools, &inventory, "legacy");

                            // Prepend tables to chapter content
                            let mut new_content = String::new();
                            if !parts_table.is_empty() {
                                new_content.push_str(&parts_table);
                                new_content.push_str("\n\n");
                            }
                            if !consumables_table.is_empty() {
                                new_content.push_str(&consumables_table);
                                new_content.push_str("\n\n");
                            }
                            if !tools_table.is_empty() {
                                new_content.push_str(&tools_table);
                                new_content.push_str("\n\n");
                            }
                            new_content.push_str(&ch.content);
                            ch.content = new_content;

                            // Accumulate for global BOM (legacy support - treat parts as fasteners)
                            accumulate_fasteners(parts, &inventory, &mut all_fasteners);
                            accumulate_consumables(consumables, &inventory, &mut all_consumables);
                            accumulate_tools(tools, &inventory, &mut all_tools);
                        }
                    }
                }
            }
        });

        // Create output directory
        create_output_directory()?;

        // Generate all output files
        generate_fasteners_file(&all_fasteners)?;
        generate_electronics_file(&all_electronics)?;
        generate_custom_parts_file(&all_custom_parts)?;
        generate_tools_file(&all_tools, &inventory)?;
        generate_consumables_file(&all_consumables, &inventory)?;

        Ok(book)
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct ChapterMetadata {
    sections: Option<std::collections::HashMap<String, SectionMetadata>>,
    // Keep legacy fields for backwards compatibility
    parts: Option<Vec<PartReference>>,
    consumables: Option<Vec<ConsumableReference>>,
    tools: Option<Vec<ToolReference>>,
}

#[derive(Debug, Deserialize, Serialize)]
struct SectionMetadata {
    hardware: Option<Vec<PartReference>>,
    electronics: Option<Vec<PartReference>>,
    custom_parts: Option<Vec<PartReference>>,
    consumables: Option<Vec<ConsumableReference>>,
    tools: Option<Vec<ToolReference>>,
    // Keep legacy fields for backward compatibility
    fasteners: Option<Vec<PartReference>>,
    parts: Option<Vec<PartReference>>,
}

// Simplified front matter structures
#[derive(Debug, Deserialize, Serialize, Clone)]
struct PartReference {
    name: String,
    quantity: u32,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ConsumableReference {
    name: String,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ToolReference {
    name: String,
    setting: Option<String>,
}

// Inventory structures
#[derive(Debug, Deserialize, Clone)]
struct InventoryFastener {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
    #[serde(rename = "Quantity", default)]
    inventory_quantity: Option<u32>, // Quantity from Excel, optional
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryElectronic {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
    #[serde(rename = "Quantity", default)]
    inventory_quantity: Option<u32>, // Quantity from Excel, optional
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryCustomPart {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
    #[serde(rename = "Quantity", default)]
    inventory_quantity: Option<u32>, // Quantity from Excel, optional
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryConsumable {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
    #[serde(rename = "Quantity", default)]
    inventory_quantity: Option<u32>, // Quantity from Excel, optional
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryTool {
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Brand", default)]
    brand: Option<String>,
}

#[derive(Debug, Clone)]
struct BomFastenerItem {
    part_number: String,
    description: String,
    supplier: String,
    total_quantity: u32,
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomElectronicItem {
    part_number: String,
    description: String,
    supplier: String,
    total_quantity: u32,
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomCustomPartItem {
    part_number: String,
    description: String,
    supplier: String,
    total_quantity: u32,
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomConsumableItem {
    part_number: String,
    description: String,
    supplier: String,
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomToolItem {
    name: String,
    brand: String,
    settings: Vec<String>, // Multiple settings from different chapters
}

fn extract_front_matter(content: &str) -> Option<String> {
    if let Some(stripped) = content.strip_prefix("---\n") {
        if let Some(end_pos) = stripped.find("\n---\n") {
            return Some(stripped[..end_pos].to_string());
        }
    }
    None
}

fn remove_front_matter(content: &str) -> String {
    if let Some(stripped) = content.strip_prefix("---\n") {
        if let Some(end_pos) = stripped.find("\n---\n") {
            return stripped[end_pos + 4..].to_string();
        }
    }
    content.to_string()
}

fn find_step_headers(content: &str) -> Vec<(String, usize)> {
    use regex::Regex;
    let re = Regex::new(r"(?i)^##+\s+Step\s+(\d+):?.*$").unwrap();

    content
        .lines()
        .enumerate()
        .filter_map(|(line_idx, line)| {
            re.captures(line).map(|caps| {
                let step_num = caps.get(1).unwrap().as_str();
                let step_key = format!("step_{}", step_num);
                (step_key, line_idx)
            })
        })
        .collect()
}

fn insert_section_tables(
    content: &str,
    sections: &std::collections::HashMap<String, SectionMetadata>,
    inventory: &Inventory,
) -> String {
    let step_headers = find_step_headers(content);
    let lines: Vec<&str> = content.lines().collect();
    let mut result = Vec::new();

    for (line_idx, line) in lines.iter().enumerate() {
        // Check if this line is a step header and add horizontal rule before it (but not the first step)
        let is_step_header = step_headers
            .iter()
            .any(|(_, header_line_idx)| line_idx == *header_line_idx);
        if is_step_header && line_idx > 0 {
            result.push("".to_string()); // Empty line
            result.push("---".to_string()); // Horizontal rule above step
            result.push("".to_string()); // Empty line
        }

        result.push(line.to_string());

        // Check if this line is a step header we need to insert tables after
        for (step_key, header_line_idx) in &step_headers {
            if line_idx == *header_line_idx {
                if let Some(section_metadata) = sections.get(step_key) {
                    // Check both hardware and fasteners for backward compatibility
                    let hardware = section_metadata.hardware.as_deref().unwrap_or_default();
                    let legacy_fasteners = section_metadata.fasteners.as_deref().unwrap_or_default();
                    let electronics = section_metadata.electronics.as_deref().unwrap_or_default();
                    let custom_parts = section_metadata.custom_parts.as_deref().unwrap_or_default();
                    let consumables = section_metadata.consumables.as_deref().unwrap_or_default();
                    let tools = section_metadata.tools.as_deref().unwrap_or_default();
                    
                    // Legacy support
                    let legacy_parts = section_metadata.parts.as_deref().unwrap_or_default();

                    let hardware_table = generate_fasteners_table(hardware, inventory, step_key);
                    let legacy_fasteners_table = generate_fasteners_table(legacy_fasteners, inventory, step_key);
                    let legacy_parts_table = generate_fasteners_table(legacy_parts, inventory, step_key);
                    let electronics_table = generate_electronics_table(electronics, inventory, step_key);
                    let custom_parts_table = generate_custom_parts_table(custom_parts, inventory, step_key);
                    let consumables_table = generate_consumables_table(consumables, inventory, step_key);
                    let tools_table = generate_tools_table(tools, inventory, step_key);

                    let has_tables = !hardware_table.is_empty()
                        || !legacy_fasteners_table.is_empty()
                        || !legacy_parts_table.is_empty()
                        || !electronics_table.is_empty()
                        || !custom_parts_table.is_empty()
                        || !consumables_table.is_empty()
                        || !tools_table.is_empty();
                    
                    if has_tables {
                        // Add Show All button before tables
                        result.push("".to_string()); // Empty line
                        result.push(generate_show_all_button(step_key));
                    }

                    if !hardware_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(hardware_table.lines().map(|s| s.to_string()));
                    }
                    if !legacy_fasteners_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(legacy_fasteners_table.lines().map(|s| s.to_string()));
                    }
                    if !legacy_parts_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(legacy_parts_table.lines().map(|s| s.to_string()));
                    }
                    if !electronics_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(electronics_table.lines().map(|s| s.to_string()));
                    }
                    if !custom_parts_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(custom_parts_table.lines().map(|s| s.to_string()));
                    }
                    if !consumables_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(consumables_table.lines().map(|s| s.to_string()));
                    }
                    if !tools_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(tools_table.lines().map(|s| s.to_string()));
                    }

                    if has_tables {
                        result.push("".to_string()); // Empty line after BOM tables
                    }
                }
                break;
            }
        }
    }

    result.join("\n")
}

fn generate_show_all_button(section_id: &str) -> String {
    format!(r#"
<button onclick="toggleAllTables('{}')" class="bom-show-all-button" style="
    background: transparent;
    color: var(--icons, #747474);
    border: 1px solid var(--icons, #747474);
    padding: 8px 16px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    margin-bottom: 10px;
    transition: all 0.2s ease;
" onmouseover="
    this.style.color='var(--icons-hover, #000000)';
    this.style.borderColor='var(--icons-hover, #000000)';
    this.style.backgroundColor='var(--theme-hover, #e6e6e6)';
" onmouseout="
    this.style.color='var(--icons, #747474)';
    this.style.borderColor='var(--icons, #747474)';
    this.style.backgroundColor='transparent';
">
    Show All
</button>

<script>
function toggleAllTables(sectionId) {{
    const button = event.target;
    const isShowing = button.textContent === 'Hide All';
    const newState = !isShowing;
    const newText = newState ? 'Hide All' : 'Show All';
    
    button.textContent = newText;
    
    // Find all details elements for this section
    const detailsElements = [
        document.getElementById('hardware-' + sectionId),
        document.getElementById('electronics-' + sectionId),
        document.getElementById('custom_parts-' + sectionId),
        document.getElementById('consumables-' + sectionId),
        document.getElementById('tools-' + sectionId)
    ].filter(el => el !== null);
    
    detailsElements.forEach(details => {{
        details.open = newState;
    }});
}}
</script>"#, section_id)
}

fn generate_fasteners_table(parts: &[PartReference], inventory: &Inventory, section_id: &str) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"hardware-{}\">\n<summary><strong>🔩 Hardware</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.fasteners.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number, part.description.as_deref().unwrap_or("-"), part_ref.quantity
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Hardware not found in inventory</td><td>{}</td></tr>\n",
                part_ref.name, part_ref.quantity
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_electronics_table(parts: &[PartReference], inventory: &Inventory, section_id: &str) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"electronics-{}\">\n<summary><strong>🔌 Electronics</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.electronics.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number, part.description.as_deref().unwrap_or("-"), part_ref.quantity
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Electronic component not found in inventory</td><td>{}</td></tr>\n",
                part_ref.name, part_ref.quantity
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_custom_parts_table(parts: &[PartReference], inventory: &Inventory, section_id: &str) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"custom_parts-{}\">\n<summary><strong>⚙️ Custom Parts</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.custom_parts.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number, part.description.as_deref().unwrap_or("-"), part_ref.quantity
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Custom part not found in inventory</td><td>{}</td></tr>\n",
                part_ref.name, part_ref.quantity
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_consumables_table(
    consumables: &[ConsumableReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
    if consumables.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"consumables-{}\">\n<summary><strong>🧪 Consumables</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th></tr>\n</thead>\n<tbody>\n", section_id));

    for consumable_ref in consumables {
        if let Some(consumable) = inventory.consumables.get(&consumable_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td></tr>\n",
                consumable.part_number, consumable.description.as_deref().unwrap_or("-")
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Consumable not found in inventory</td></tr>\n",
                consumable_ref.name
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_tools_table(tools: &[ToolReference], inventory: &Inventory, section_id: &str) -> String {
    if tools.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"tools-{}\">\n<summary><strong>🔧 Tools</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Setting</th><th>Brand</th></tr>\n</thead>\n<tbody>\n", section_id));

    for tool_ref in tools {
        if let Some(tool) = inventory.tools.get(&tool_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                tool.name,
                tool_ref.setting.as_deref().unwrap_or("-"),
                tool.brand.as_deref().unwrap_or("-")
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>Tool not found in inventory</td></tr>\n",
                tool_ref.name,
                tool_ref.setting.as_deref().unwrap_or("-")
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn accumulate_fasteners(
    parts: &[PartReference],
    inventory: &Inventory,
    all_fasteners: &mut HashMap<String, BomFastenerItem>,
) {
    for part_ref in parts {
        if let Some(inventory_part) = inventory.fasteners.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_fasteners
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomFastenerItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part.description.as_deref().unwrap_or("-").to_string(),
                    supplier: "N/A".to_string(), // No supplier in Excel
                    total_quantity: part_ref.quantity,
                    unit_cost: None, // No unit cost in Excel
                });
        }
    }
}

fn accumulate_electronics(
    parts: &[PartReference],
    inventory: &Inventory,
    all_electronics: &mut HashMap<String, BomElectronicItem>,
) {
    for part_ref in parts {
        if let Some(inventory_part) = inventory.electronics.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_electronics
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomElectronicItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part.description.as_deref().unwrap_or("-").to_string(),
                    supplier: "N/A".to_string(), // No supplier in Excel
                    total_quantity: part_ref.quantity,
                    unit_cost: None, // No unit cost in Excel
                });
        }
    }
}

fn accumulate_custom_parts(
    parts: &[PartReference],
    inventory: &Inventory,
    all_custom_parts: &mut HashMap<String, BomCustomPartItem>,
) {
    for part_ref in parts {
        if let Some(inventory_part) = inventory.custom_parts.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_custom_parts
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomCustomPartItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part.description.as_deref().unwrap_or("-").to_string(),
                    supplier: "N/A".to_string(), // No supplier in Excel
                    total_quantity: part_ref.quantity,
                    unit_cost: None, // No unit cost in Excel
                });
        }
    }
}

fn accumulate_consumables(
    consumables: &[ConsumableReference],
    inventory: &Inventory,
    all_consumables: &mut HashMap<String, BomConsumableItem>,
) {
    for consumable_ref in consumables {
        if let Some(inventory_consumable) = inventory.consumables.get(&consumable_ref.name) {
            let key = consumable_ref.name.clone();

            // For consumables, we'll just track unique items (not quantities since they're often descriptive)
            all_consumables
                .entry(key)
                .or_insert_with(|| BomConsumableItem {
                    part_number: inventory_consumable.part_number.clone(),
                    description: inventory_consumable.description.as_deref().unwrap_or("-").to_string(),
                    supplier: "N/A".to_string(), // No supplier in Excel
                    unit_cost: None, // No unit cost in Excel
                });
        }
    }
}

fn accumulate_tools(
    tools: &[ToolReference],
    inventory: &Inventory,
    all_tools: &mut HashMap<String, BomToolItem>,
) {
    for tool_ref in tools {
        if let Some(inventory_tool) = inventory.tools.get(&tool_ref.name) {
            let key = tool_ref.name.clone();

            all_tools
                .entry(key)
                .and_modify(|item| {
                    if let Some(setting) = &tool_ref.setting {
                        if !item.settings.contains(setting) {
                            item.settings.push(setting.clone());
                        }
                    }
                })
                .or_insert_with(|| {
                    let mut settings = Vec::new();
                    if let Some(setting) = &tool_ref.setting {
                        settings.push(setting.clone());
                    }
                    BomToolItem {
                        name: inventory_tool.name.clone(),
                        brand: inventory_tool.brand.as_deref().unwrap_or("-").to_string(),
                        settings,
                    }
                });
        }
    }
}

fn create_output_directory() -> Result<(), Error> {
    std::fs::create_dir_all("output")
        .map_err(|e| Error::msg(format!("Failed to create output directory: {}", e)))?;
    Ok(())
}

fn generate_fasteners_file(fasteners: &HashMap<String, BomFastenerItem>) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Part Number,Description,Quantity\n");

    // Fasteners section
    let mut sorted_fasteners: Vec<_> = fasteners.values().collect();
    sorted_fasteners.sort_by(|a, b| a.description.cmp(&b.description));

    for fastener in sorted_fasteners {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\",{}\n",
            fastener.part_number,
            fastener.description,
            fastener.total_quantity
        ));
    }

    // Write fasteners to CSV file
    std::fs::write("output/hardware.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write hardware CSV file: {}", e)))?;

    Ok(())
}

fn generate_electronics_file(electronics: &HashMap<String, BomElectronicItem>) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Name,Description,Quantity\n");

    // Electronics section
    let mut sorted_electronics: Vec<_> = electronics.values().collect();
    sorted_electronics.sort_by(|a, b| a.description.cmp(&b.description));

    for electronic in sorted_electronics {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\",{}\n",
            electronic.part_number,
            electronic.description,
            electronic.total_quantity
        ));
    }

    // Write electronics to CSV file
    std::fs::write("output/electronics.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write electronics CSV file: {}", e)))?;

    Ok(())
}

fn generate_custom_parts_file(custom_parts: &HashMap<String, BomCustomPartItem>) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Name,Description,Quantity\n");

    // Custom parts section
    let mut sorted_custom_parts: Vec<_> = custom_parts.values().collect();
    sorted_custom_parts.sort_by(|a, b| a.description.cmp(&b.description));

    for custom_part in sorted_custom_parts {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\",{}\n",
            custom_part.part_number,
            custom_part.description,
            custom_part.total_quantity
        ));
    }

    // Write custom parts to CSV file
    std::fs::write("output/custom_parts.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write custom parts CSV file: {}", e)))?;

    Ok(())
}

fn generate_tools_file(
    tools: &HashMap<String, BomToolItem>,
    _inventory: &Inventory,
) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Name,Brand\n");

    // Tools section - only include tools that were actually used
    let mut sorted_tools: Vec<_> = tools.values().collect();
    sorted_tools.sort_by(|a, b| a.name.cmp(&b.name));

    for tool in sorted_tools {
        csv_content.push_str(&format!("\"{}\",\"{}\"\n", tool.name, tool.brand));
    }

    // Write tools to CSV file
    std::fs::write("output/tools.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write tools CSV file: {}", e)))?;

    Ok(())
}

fn generate_consumables_file(
    consumables: &HashMap<String, BomConsumableItem>,
    _inventory: &Inventory,
) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Name,Description\n");

    // Consumables section - only include consumables that were actually used
    let mut sorted_consumables: Vec<_> = consumables.values().collect();
    sorted_consumables.sort_by(|a, b| a.description.cmp(&b.description));

    for consumable in sorted_consumables {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\"\n",
            consumable.part_number,
            consumable.description
        ));
    }

    // Write consumables to CSV file
    std::fs::write("output/consumables.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write consumables CSV file: {}", e)))?;

    Ok(())
}
