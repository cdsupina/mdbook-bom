use calamine::{open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use clap::{Arg, ArgMatches, Command};
use mdbook::book::{Book, BookItem};
use mdbook::errors::Error;
use mdbook::preprocess::{CmdPreprocessor, Preprocessor, PreprocessorContext};
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io;

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
    // Load .env file if present (for local configuration)
    // Silently ignore if .env file doesn't exist
    let _ = dotenvy::dotenv();

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
    assemblies: HashMap<String, InventoryAssembly>,
}

impl Inventory {
    fn load(excel_path: &str) -> Result<Self, Error> {
        // Expand home directory if needed
        let expanded_path = if let Some(stripped) = excel_path.strip_prefix("~/") {
            if let Some(home) = std::env::var_os("HOME") {
                let home_path = std::path::Path::new(&home);
                home_path.join(stripped).to_string_lossy().to_string()
            } else {
                return Err(Error::msg(
                    "Cannot expand ~ - HOME environment variable not set",
                ));
            }
        } else {
            excel_path.to_string()
        };

        // Check if file exists first
        if !std::path::Path::new(&expanded_path).exists() {
            return Err(Error::msg(format!(
                "Excel file not found: {}",
                expanded_path
            )));
        }

        let fasteners = Self::load_fasteners_from_excel(&expanded_path)?;
        let electronics = Self::load_electronics_from_excel(&expanded_path)?;
        let custom_parts = Self::load_custom_parts_from_excel(&expanded_path)?;
        let consumables = Self::load_consumables_from_excel(&expanded_path)?;
        let tools = Self::load_tools_from_excel(&expanded_path)?;
        let assemblies = Self::load_assemblies_from_excel(&expanded_path)?;

        Ok(Inventory {
            fasteners,
            electronics,
            custom_parts,
            consumables,
            tools,
            assemblies,
        })
    }

    fn load_fasteners_from_excel(
        excel_path: &str,
    ) -> Result<HashMap<String, InventoryFastener>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Hardware")
            .map_err(|e| Error::msg(format!("Failed to read 'Hardware' sheet: {}", e)))?;

        let mut hardware = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| {
                Error::msg(format!("Failed to create deserializer for hardware: {}", e))
            })?;

        for result in iter {
            let hardware_item: InventoryFastener =
                result.map_err(|e| Error::msg(format!("Failed to parse hardware row: {}", e)))?;
            hardware.insert(hardware_item.part_number.clone(), hardware_item);
        }

        Ok(hardware)
    }

    fn load_electronics_from_excel(
        excel_path: &str,
    ) -> Result<HashMap<String, InventoryElectronic>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Electronics")
            .map_err(|e| Error::msg(format!("Failed to read 'Electronics' sheet: {}", e)))?;

        let mut electronics = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| {
                Error::msg(format!(
                    "Failed to create deserializer for electronics: {}",
                    e
                ))
            })?;

        for result in iter {
            let electronic: InventoryElectronic =
                result.map_err(|e| Error::msg(format!("Failed to parse electronic row: {}", e)))?;
            electronics.insert(electronic.part_number.clone(), electronic);
        }

        Ok(electronics)
    }

    fn load_custom_parts_from_excel(
        excel_path: &str,
    ) -> Result<HashMap<String, InventoryCustomPart>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Custom Parts")
            .map_err(|e| Error::msg(format!("Failed to read 'Custom Parts' sheet: {}", e)))?;

        let mut custom_parts = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| {
                Error::msg(format!(
                    "Failed to create deserializer for custom_parts: {}",
                    e
                ))
            })?;

        for result in iter {
            let custom_part: InventoryCustomPart = result
                .map_err(|e| Error::msg(format!("Failed to parse custom_part row: {}", e)))?;
            custom_parts.insert(custom_part.part_number.clone(), custom_part);
        }

        Ok(custom_parts)
    }

    fn load_consumables_from_excel(
        excel_path: &str,
    ) -> Result<HashMap<String, InventoryConsumable>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Consumables")
            .map_err(|e| Error::msg(format!("Failed to read 'Consumables' sheet: {}", e)))?;

        let mut consumables = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| {
                Error::msg(format!(
                    "Failed to create deserializer for consumables: {}",
                    e
                ))
            })?;

        for result in iter {
            let consumable: InventoryConsumable =
                result.map_err(|e| Error::msg(format!("Failed to parse consumable row: {}", e)))?;
            consumables.insert(consumable.part_number.clone(), consumable);
        }

        Ok(consumables)
    }

    fn load_tools_from_excel(excel_path: &str) -> Result<HashMap<String, InventoryTool>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Tools")
            .map_err(|e| Error::msg(format!("Failed to read 'Tools' sheet: {}", e)))?;

        let mut tools = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| Error::msg(format!("Failed to create deserializer for tools: {}", e)))?;

        for result in iter {
            let tool: InventoryTool =
                result.map_err(|e| Error::msg(format!("Failed to parse tool row: {}", e)))?;
            tools.insert(tool.name.clone(), tool);
        }

        Ok(tools)
    }

    fn load_assemblies_from_excel(
        excel_path: &str,
    ) -> Result<HashMap<String, InventoryAssembly>, Error> {
        let mut workbook: Xlsx<_> = open_workbook(excel_path)
            .map_err(|e| Error::msg(format!("Failed to open Excel file: {}", e)))?;

        let range = workbook
            .worksheet_range("Assemblies")
            .map_err(|e| Error::msg(format!("Failed to read 'Assemblies' sheet: {}", e)))?;

        let mut assemblies = HashMap::new();
        let iter = RangeDeserializerBuilder::new()
            .from_range(&range)
            .map_err(|e| {
                Error::msg(format!(
                    "Failed to create deserializer for assemblies: {}",
                    e
                ))
            })?;

        for result in iter {
            let assembly: InventoryAssembly =
                result.map_err(|e| Error::msg(format!("Failed to parse assembly row: {}", e)))?;
            assemblies.insert(assembly.name.clone(), assembly);
        }

        Ok(assemblies)
    }
}

pub struct BomPreprocessor;

impl Preprocessor for BomPreprocessor {
    fn name(&self) -> &str {
        "bom"
    }

    fn run(&self, _ctx: &PreprocessorContext, mut book: Book) -> Result<Book, Error> {
        // Read configuration from environment variables (loaded from .env file)
        let excel_path = std::env::var("BOM_INVENTORY_FILE")
            .map_err(|_| Error::msg("BOM_INVENTORY_FILE environment variable is required. Set it in .env file in the book directory."))?;

        let output_path = std::env::var("BOM_OUTPUT_PATH")
            .map_err(|_| Error::msg("BOM_OUTPUT_PATH environment variable is required. Set it in .env file in the book directory."))?;

        // Load inventory data
        let inventory = Inventory::load(&excel_path)?;

        let mut all_fasteners: HashMap<String, BomFastenerItem> = HashMap::new();
        let mut all_electronics: HashMap<String, BomElectronicItem> = HashMap::new();
        let mut all_custom_parts: HashMap<String, BomCustomPartItem> = HashMap::new();
        let mut all_consumables: HashMap<String, BomConsumableItem> = HashMap::new();
        let mut all_tools: HashMap<String, BomToolItem> = HashMap::new();
        let mut all_assemblies: HashMap<String, BomAssemblyItem> = HashMap::new();

        book.for_each_mut(|item: &mut BookItem| {
            if let BookItem::Chapter(ch) = item {
                if let Some(front_matter) = extract_front_matter(&ch.content) {
                    // Remove front matter from content
                    let content_without_fm = remove_front_matter(&ch.content);

                    // Parse YAML
                    if let Ok(metadata) = serde_yml::from_str::<ChapterMetadata>(&front_matter) {
                        // Insert tables after step headers
                        ch.content =
                            insert_section_tables(&content_without_fm, &metadata.sections, &inventory);

                        // Accumulate all items from all sections for BOM
                        for section_metadata in metadata.sections.values() {
                            let hardware =
                                section_metadata.hardware.as_deref().unwrap_or_default();
                            let electronics =
                                section_metadata.electronics.as_deref().unwrap_or_default();
                            let custom_parts =
                                section_metadata.custom_parts.as_deref().unwrap_or_default();
                            let consumables =
                                section_metadata.consumables.as_deref().unwrap_or_default();
                            let tools = section_metadata.tools.as_deref().unwrap_or_default();
                            let assemblies =
                                section_metadata.assemblies.as_deref().unwrap_or_default();

                            accumulate_fasteners(hardware, &inventory, &mut all_fasteners);
                            accumulate_electronics(
                                electronics,
                                &inventory,
                                &mut all_electronics,
                            );
                            accumulate_custom_parts(
                                custom_parts,
                                &inventory,
                                &mut all_custom_parts,
                            );
                            accumulate_consumables(
                                consumables,
                                &inventory,
                                &mut all_consumables,
                            );
                            accumulate_tools(tools, &inventory, &mut all_tools);
                            accumulate_assemblies(
                                assemblies,
                                &inventory,
                                &mut all_assemblies,
                            );
                        }
                    }
                }
            }
        });

        // Create directory for output file
        create_output_directory_for_path(&output_path)?;

        // Generate BOM Excel file
        generate_bom_excel_file(
            &all_fasteners,
            &all_electronics,
            &all_custom_parts,
            &all_consumables,
            &all_tools,
            &all_assemblies,
            &output_path,
        )?;

        Ok(book)
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct ChapterMetadata {
    sections: std::collections::HashMap<String, SectionMetadata>,
}

#[derive(Debug, Deserialize, Serialize)]
struct SectionMetadata {
    hardware: Option<Vec<PartReference>>,
    electronics: Option<Vec<PartReference>>,
    custom_parts: Option<Vec<PartReference>>,
    consumables: Option<Vec<ConsumableReference>>,
    tools: Option<Vec<ToolReference>>,
    assemblies: Option<Vec<AssemblyReference>>,
}

// Simplified front matter structures
#[derive(Debug, Deserialize, Serialize, Clone)]
struct PartReference {
    name: String,
    quantity: u32,
    #[serde(default)]
    exclude_from_bom: bool,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ConsumableReference {
    name: String,
    #[serde(default)]
    exclude_from_bom: bool,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ToolReference {
    name: String,
    setting: Option<String>,
    #[serde(default)]
    exclude_from_bom: bool,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct AssemblyReference {
    name: String,
    quantity: u32,
    #[serde(default)]
    exclude_from_bom: bool,
}

// Inventory structures
#[derive(Debug, Deserialize, Clone)]
struct InventoryFastener {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryElectronic {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryCustomPart {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryConsumable {
    #[serde(rename = "Name")]
    part_number: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryTool {
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Brand", default)]
    brand: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryAssembly {
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Description", default)]
    description: Option<String>,
}

#[derive(Debug, Clone)]
struct BomFastenerItem {
    part_number: String,
    description: String,
    #[allow(dead_code)]
    supplier: String,
    total_quantity: u32,
    #[allow(dead_code)]
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomElectronicItem {
    part_number: String,
    description: String,
    #[allow(dead_code)]
    supplier: String,
    total_quantity: u32,
    #[allow(dead_code)]
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomCustomPartItem {
    part_number: String,
    description: String,
    #[allow(dead_code)]
    supplier: String,
    total_quantity: u32,
    #[allow(dead_code)]
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomConsumableItem {
    part_number: String,
    description: String,
    #[allow(dead_code)]
    supplier: String,
    #[allow(dead_code)]
    unit_cost: Option<f64>,
}

#[derive(Debug, Clone)]
struct BomToolItem {
    name: String,
    brand: String,
    settings: Vec<String>, // Multiple settings from different chapters
}

#[derive(Debug, Clone)]
struct BomAssemblyItem {
    name: String,
    description: String,
    total_quantity: u32,
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
    let mut overview_inserted = false;

    // Generate overview tables (without header)
    let overview_section = generate_overview_tables(sections, inventory);

    for (line_idx, line) in lines.iter().enumerate() {
        // Check if this is a top-level header (# Header) and insert overview after it
        if !overview_inserted && line.starts_with("# ") && !line.starts_with("## ") {
            result.push(line.to_string());

            // Insert overview tables after the top-level header
            if !overview_section.trim().is_empty() {
                result.push("".to_string()); // Empty line
                result.extend(overview_section.lines().map(|s| s.to_string()));
                result.push("".to_string()); // Empty line after overview
            }
            overview_inserted = true;
            continue;
        }

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
                    let hardware = section_metadata.hardware.as_deref().unwrap_or_default();
                    let electronics = section_metadata.electronics.as_deref().unwrap_or_default();
                    let custom_parts = section_metadata.custom_parts.as_deref().unwrap_or_default();
                    let consumables = section_metadata.consumables.as_deref().unwrap_or_default();
                    let tools = section_metadata.tools.as_deref().unwrap_or_default();
                    let assemblies = section_metadata.assemblies.as_deref().unwrap_or_default();

                    let hardware_table = generate_fasteners_table(hardware, inventory, step_key);
                    let electronics_table =
                        generate_electronics_table(electronics, inventory, step_key);
                    let custom_parts_table =
                        generate_custom_parts_table(custom_parts, inventory, step_key);
                    let consumables_table =
                        generate_consumables_table(consumables, inventory, step_key);
                    let tools_table = generate_tools_table(tools, inventory, step_key);
                    let assemblies_table =
                        generate_assemblies_table(assemblies, inventory, step_key);

                    let has_tables = !hardware_table.is_empty()
                        || !electronics_table.is_empty()
                        || !custom_parts_table.is_empty()
                        || !consumables_table.is_empty()
                        || !tools_table.is_empty()
                        || !assemblies_table.is_empty();

                    if has_tables {
                        // Add Show All button before tables
                        result.push("".to_string()); // Empty line
                        result.push(generate_show_all_button(step_key));
                    }

                    if !hardware_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(hardware_table.lines().map(|s| s.to_string()));
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
                    if !assemblies_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(assemblies_table.lines().map(|s| s.to_string()));
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

fn generate_overview_tables(
    sections: &std::collections::HashMap<String, SectionMetadata>,
    inventory: &Inventory,
) -> String {
    // Aggregate all parts from all sections
    let mut all_hardware = Vec::new();
    let mut all_electronics = Vec::new();
    let mut all_custom_parts = Vec::new();
    let mut all_consumables = Vec::new();
    let mut all_tools = Vec::new();
    let mut all_assemblies = Vec::new();

    for section_metadata in sections.values() {
        // Collect hardware
        if let Some(hardware) = &section_metadata.hardware {
            all_hardware.extend(hardware.clone());
        }

        // Collect other categories
        if let Some(electronics) = &section_metadata.electronics {
            all_electronics.extend(electronics.clone());
        }
        if let Some(custom_parts) = &section_metadata.custom_parts {
            all_custom_parts.extend(custom_parts.clone());
        }
        if let Some(consumables) = &section_metadata.consumables {
            all_consumables.extend(consumables.clone());
        }
        if let Some(tools) = &section_metadata.tools {
            all_tools.extend(tools.clone());
        }
        if let Some(assemblies) = &section_metadata.assemblies {
            all_assemblies.extend(assemblies.clone());
        }
    }

    // Deduplicate and combine quantities
    let combined_hardware = combine_parts(&all_hardware);
    let combined_electronics = combine_parts(&all_electronics);
    let combined_custom_parts = combine_parts(&all_custom_parts);
    let combined_consumables = deduplicate_consumables(&all_consumables);
    let combined_tools = deduplicate_tools(&all_tools);
    let combined_assemblies = combine_assemblies(&all_assemblies);

    let mut overview = String::new();

    // Generate overview tables
    let hardware_table = generate_fasteners_table(&combined_hardware, inventory, "overview");
    let electronics_table =
        generate_electronics_table(&combined_electronics, inventory, "overview");
    let custom_parts_table =
        generate_custom_parts_table(&combined_custom_parts, inventory, "overview");
    let consumables_table =
        generate_consumables_table(&combined_consumables, inventory, "overview");
    let tools_table = generate_tools_table(&combined_tools, inventory, "overview");
    let assemblies_table =
        generate_assemblies_table(&combined_assemblies, inventory, "overview");

    let has_tables = !hardware_table.is_empty()
        || !electronics_table.is_empty()
        || !custom_parts_table.is_empty()
        || !consumables_table.is_empty()
        || !tools_table.is_empty()
        || !assemblies_table.is_empty();

    if has_tables {
        overview.push_str(&generate_show_all_button("overview"));
        overview.push('\n');

        if !hardware_table.is_empty() {
            overview.push_str(&hardware_table);
            overview.push('\n');
        }
        if !electronics_table.is_empty() {
            overview.push_str(&electronics_table);
            overview.push('\n');
        }
        if !custom_parts_table.is_empty() {
            overview.push_str(&custom_parts_table);
            overview.push('\n');
        }
        if !consumables_table.is_empty() {
            overview.push_str(&consumables_table);
            overview.push('\n');
        }
        if !tools_table.is_empty() {
            overview.push_str(&tools_table);
            overview.push('\n');
        }
        if !assemblies_table.is_empty() {
            overview.push_str(&assemblies_table);
            overview.push('\n');
        }
    }

    overview
}

fn combine_parts(parts: &[PartReference]) -> Vec<PartReference> {
    let mut combined: std::collections::HashMap<String, (u32, bool)> =
        std::collections::HashMap::new();

    for part in parts {
        combined
            .entry(part.name.clone())
            .and_modify(|(qty, excluded)| {
                *qty += part.quantity;
                // Only exclude if ALL references exclude it
                *excluded = *excluded && part.exclude_from_bom;
            })
            .or_insert((part.quantity, part.exclude_from_bom));
    }

    combined
        .into_iter()
        .map(|(name, (quantity, exclude_from_bom))| PartReference {
            name,
            quantity,
            exclude_from_bom,
        })
        .collect()
}

fn deduplicate_consumables(consumables: &[ConsumableReference]) -> Vec<ConsumableReference> {
    let mut combined: std::collections::HashMap<String, bool> =
        std::collections::HashMap::new();

    for consumable in consumables {
        combined
            .entry(consumable.name.clone())
            .and_modify(|excluded| {
                // Only exclude if ALL references exclude it
                *excluded = *excluded && consumable.exclude_from_bom;
            })
            .or_insert(consumable.exclude_from_bom);
    }

    combined
        .into_iter()
        .map(|(name, exclude_from_bom)| ConsumableReference {
            name,
            exclude_from_bom,
        })
        .collect()
}

fn deduplicate_tools(tools: &[ToolReference]) -> Vec<ToolReference> {
    let mut combined: std::collections::HashMap<
        String,
        (std::collections::HashSet<String>, bool),
    > = std::collections::HashMap::new();

    for tool in tools {
        let entry = combined
            .entry(tool.name.clone())
            .or_insert_with(|| (std::collections::HashSet::new(), tool.exclude_from_bom));
        if let Some(setting) = &tool.setting {
            entry.0.insert(setting.clone());
        }
        // Only exclude if ALL references exclude it
        entry.1 = entry.1 && tool.exclude_from_bom;
    }

    combined
        .into_iter()
        .map(|(name, (settings, exclude_from_bom))| {
            let setting = if settings.is_empty() {
                None
            } else {
                Some(settings.into_iter().collect::<Vec<_>>().join(", "))
            };
            ToolReference {
                name,
                setting,
                exclude_from_bom,
            }
        })
        .collect()
}

fn generate_show_all_button(section_id: &str) -> String {
    format!(
        r#"
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
        document.getElementById('tools-' + sectionId),
        document.getElementById('assemblies-' + sectionId)
    ].filter(el => el !== null);

    detailsElements.forEach(details => {{
        details.open = newState;
    }});
}}
</script>"#,
        section_id
    )
}

fn generate_fasteners_table(
    parts: &[PartReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"hardware-{}\">\n<summary><strong>🔩 Hardware</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.fasteners.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number,
                part.description.as_deref().unwrap_or("-"),
                part_ref.quantity
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

fn generate_electronics_table(
    parts: &[PartReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"electronics-{}\">\n<summary><strong>🔌 Electronics</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.electronics.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number,
                part.description.as_deref().unwrap_or("-"),
                part_ref.quantity
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

fn generate_custom_parts_table(
    parts: &[PartReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"custom_parts-{}\">\n<summary><strong>⚙️ Custom Parts</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for part_ref in parts {
        if let Some(part) = inventory.custom_parts.get(&part_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number,
                part.description.as_deref().unwrap_or("-"),
                part_ref.quantity
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
                consumable.part_number,
                consumable.description.as_deref().unwrap_or("-")
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

fn generate_tools_table(
    tools: &[ToolReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
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

fn generate_assemblies_table(
    assemblies: &[AssemblyReference],
    inventory: &Inventory,
    section_id: &str,
) -> String {
    if assemblies.is_empty() {
        return String::new();
    }

    let mut table = String::from(&format!("<details id=\"assemblies-{}\">\n<summary><strong>\u{1f9e9} Assemblies</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Description</th><th>Quantity</th></tr>\n</thead>\n<tbody>\n", section_id));

    for assembly_ref in assemblies {
        if let Some(assembly) = inventory.assemblies.get(&assembly_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                assembly.name,
                assembly.description.as_deref().unwrap_or("-"),
                assembly_ref.quantity
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Assembly not found in inventory</td><td>{}</td></tr>\n",
                assembly_ref.name, assembly_ref.quantity
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn combine_assemblies(assemblies: &[AssemblyReference]) -> Vec<AssemblyReference> {
    let mut combined: std::collections::HashMap<String, (u32, bool)> =
        std::collections::HashMap::new();

    for assembly in assemblies {
        combined
            .entry(assembly.name.clone())
            .and_modify(|(qty, excluded)| {
                *qty += assembly.quantity;
                // Only exclude if ALL references exclude it
                *excluded = *excluded && assembly.exclude_from_bom;
            })
            .or_insert((assembly.quantity, assembly.exclude_from_bom));
    }

    combined
        .into_iter()
        .map(|(name, (quantity, exclude_from_bom))| AssemblyReference {
            name,
            quantity,
            exclude_from_bom,
        })
        .collect()
}

fn accumulate_assemblies(
    assemblies: &[AssemblyReference],
    inventory: &Inventory,
    all_assemblies: &mut HashMap<String, BomAssemblyItem>,
) {
    for assembly_ref in assemblies {
        if assembly_ref.exclude_from_bom {
            continue;
        }

        if let Some(inventory_assembly) = inventory.assemblies.get(&assembly_ref.name) {
            let key = assembly_ref.name.clone();

            all_assemblies
                .entry(key)
                .and_modify(|item| item.total_quantity += assembly_ref.quantity)
                .or_insert_with(|| BomAssemblyItem {
                    name: inventory_assembly.name.clone(),
                    description: inventory_assembly
                        .description
                        .as_deref()
                        .unwrap_or("-")
                        .to_string(),
                    total_quantity: assembly_ref.quantity,
                });
        }
    }
}

fn accumulate_fasteners(
    parts: &[PartReference],
    inventory: &Inventory,
    all_fasteners: &mut HashMap<String, BomFastenerItem>,
) {
    for part_ref in parts {
        if part_ref.exclude_from_bom {
            continue;
        }
        if let Some(inventory_part) = inventory.fasteners.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_fasteners
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomFastenerItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part
                        .description
                        .as_deref()
                        .unwrap_or("-")
                        .to_string(),
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
        if part_ref.exclude_from_bom {
            continue;
        }
        if let Some(inventory_part) = inventory.electronics.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_electronics
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomElectronicItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part
                        .description
                        .as_deref()
                        .unwrap_or("-")
                        .to_string(),
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
        if part_ref.exclude_from_bom {
            continue;
        }
        if let Some(inventory_part) = inventory.custom_parts.get(&part_ref.name) {
            let key = part_ref.name.clone();

            all_custom_parts
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomCustomPartItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part
                        .description
                        .as_deref()
                        .unwrap_or("-")
                        .to_string(),
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
        if consumable_ref.exclude_from_bom {
            continue;
        }
        if let Some(inventory_consumable) = inventory.consumables.get(&consumable_ref.name) {
            let key = consumable_ref.name.clone();

            // For consumables, we'll just track unique items (not quantities since they're often descriptive)
            all_consumables
                .entry(key)
                .or_insert_with(|| BomConsumableItem {
                    part_number: inventory_consumable.part_number.clone(),
                    description: inventory_consumable
                        .description
                        .as_deref()
                        .unwrap_or("-")
                        .to_string(),
                    supplier: "N/A".to_string(), // No supplier in Excel
                    unit_cost: None,             // No unit cost in Excel
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
        if tool_ref.exclude_from_bom {
            continue;
        }
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

fn create_output_directory_for_path(file_path: &str) -> Result<(), Error> {
    if let Some(parent_dir) = std::path::Path::new(file_path).parent() {
        std::fs::create_dir_all(parent_dir).map_err(|e| {
            Error::msg(format!(
                "Failed to create directory '{}': {}",
                parent_dir.display(),
                e
            ))
        })?;
    }
    Ok(())
}

fn generate_bom_excel_file(
    fasteners: &HashMap<String, BomFastenerItem>,
    electronics: &HashMap<String, BomElectronicItem>,
    custom_parts: &HashMap<String, BomCustomPartItem>,
    consumables: &HashMap<String, BomConsumableItem>,
    tools: &HashMap<String, BomToolItem>,
    assemblies: &HashMap<String, BomAssemblyItem>,
    output_path: &str,
) -> Result<(), Error> {
    let mut workbook = Workbook::new();

    // Generate Hardware sheet
    if !fasteners.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Hardware")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Part Number")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Description")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 2, "Quantity")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_fasteners: Vec<_> = fasteners.values().collect();
        sorted_fasteners.sort_by(|a, b| a.description.cmp(&b.description));

        for (row, fastener) in sorted_fasteners.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &fastener.part_number)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &fastener.description)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_number(row as u32, 2, fastener.total_quantity as f64)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    // Generate Electronics sheet
    if !electronics.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Electronics")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Name")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Description")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 2, "Quantity")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_electronics: Vec<_> = electronics.values().collect();
        sorted_electronics.sort_by(|a, b| a.description.cmp(&b.description));

        for (row, electronic) in sorted_electronics.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &electronic.part_number)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &electronic.description)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_number(row as u32, 2, electronic.total_quantity as f64)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    // Generate Custom Parts sheet
    if !custom_parts.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Custom Parts")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Name")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Description")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 2, "Quantity")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_custom_parts: Vec<_> = custom_parts.values().collect();
        sorted_custom_parts.sort_by(|a, b| a.description.cmp(&b.description));

        for (row, custom_part) in sorted_custom_parts.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &custom_part.part_number)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &custom_part.description)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_number(row as u32, 2, custom_part.total_quantity as f64)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    // Generate Tools sheet
    if !tools.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Tools")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Name")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Brand")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_tools: Vec<_> = tools.values().collect();
        sorted_tools.sort_by(|a, b| a.name.cmp(&b.name));

        for (row, tool) in sorted_tools.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &tool.name)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &tool.brand)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    // Generate Consumables sheet
    if !consumables.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Consumables")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Name")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Description")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_consumables: Vec<_> = consumables.values().collect();
        sorted_consumables.sort_by(|a, b| a.description.cmp(&b.description));

        for (row, consumable) in sorted_consumables.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &consumable.part_number)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &consumable.description)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    // Generate Assemblies sheet
    if !assemblies.is_empty() {
        let worksheet = workbook
            .add_worksheet()
            .set_name("Assemblies")
            .map_err(|e| Error::msg(format!("Failed to set sheet name: {}", e)))?;

        // Headers
        worksheet
            .write_string(0, 0, "Name")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 1, "Description")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;
        worksheet
            .write_string(0, 2, "Quantity")
            .map_err(|e| Error::msg(format!("Failed to write header: {}", e)))?;

        // Data
        let mut sorted_assemblies: Vec<_> = assemblies.values().collect();
        sorted_assemblies.sort_by(|a, b| a.name.cmp(&b.name));

        for (row, assembly) in sorted_assemblies.iter().enumerate() {
            let row = row + 1; // Skip header row
            worksheet
                .write_string(row as u32, 0, &assembly.name)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_string(row as u32, 1, &assembly.description)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
            worksheet
                .write_number(row as u32, 2, assembly.total_quantity as f64)
                .map_err(|e| Error::msg(format!("Failed to write data: {}", e)))?;
        }
    }

    workbook
        .save(output_path)
        .map_err(|e| Error::msg(format!("Failed to save Excel file: {}", e)))?;

    Ok(())
}
