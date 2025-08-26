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
    parts: HashMap<String, InventoryPart>,
    consumables: HashMap<String, InventoryConsumable>,
    tools: HashMap<String, InventoryTool>,
}

impl Inventory {
    fn load() -> Result<Self, Error> {
        let parts = Self::load_parts()?;
        let consumables = Self::load_consumables()?;
        let tools = Self::load_tools()?;

        Ok(Inventory {
            parts,
            consumables,
            tools,
        })
    }

    fn load_parts() -> Result<HashMap<String, InventoryPart>, Error> {
        let path = Path::new("inventory/parts.csv");
        if !path.exists() {
            return Err(Error::msg("inventory/parts.csv not found"));
        }

        let mut reader = csv::Reader::from_path(path)
            .map_err(|e| Error::msg(format!("Failed to read parts.csv: {}", e)))?;

        let mut parts = HashMap::new();
        for result in reader.deserialize() {
            let part: InventoryPart =
                result.map_err(|e| Error::msg(format!("Failed to parse part: {}", e)))?;
            parts.insert(part.part_number.clone(), part);
        }

        Ok(parts)
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
}

pub struct BomPreprocessor;

impl Preprocessor for BomPreprocessor {
    fn name(&self) -> &str {
        "mdbook-bom"
    }

    fn run(&self, _ctx: &PreprocessorContext, mut book: Book) -> Result<Book, Error> {
        // Load inventory data
        let inventory = Inventory::load()?;

        let mut all_parts: HashMap<String, BomItem> = HashMap::new();
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
                                let parts = section_metadata.parts.as_deref().unwrap_or_default();
                                let consumables =
                                    section_metadata.consumables.as_deref().unwrap_or_default();
                                let tools = section_metadata.tools.as_deref().unwrap_or_default();

                                accumulate_parts(parts, &inventory, &mut all_parts);
                                accumulate_consumables(
                                    consumables,
                                    &inventory,
                                    &mut all_consumables,
                                );
                                accumulate_tools(tools, &inventory, &mut all_tools);
                            }
                        } else {
                            // Handle legacy flat structure (backwards compatibility)
                            ch.content = content_without_fm;

                            let parts = metadata.parts.as_deref().unwrap_or_default();
                            let consumables = metadata.consumables.as_deref().unwrap_or_default();
                            let tools = metadata.tools.as_deref().unwrap_or_default();

                            // Generate tables for this chapter (legacy behavior)
                            let parts_table = generate_parts_table(parts, &inventory);
                            let consumables_table =
                                generate_consumables_table(consumables, &inventory);
                            let tools_table = generate_tools_table(tools, &inventory);

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

                            // Accumulate for global BOM
                            accumulate_parts(parts, &inventory, &mut all_parts);
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
        generate_bom_file(&all_parts)?;
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
    parts: Option<Vec<PartReference>>,
    consumables: Option<Vec<ConsumableReference>>,
    tools: Option<Vec<ToolReference>>,
}

// Simplified front matter structures
#[derive(Debug, Deserialize, Serialize, Clone)]
struct PartReference {
    part_number: String,
    quantity: u32,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ConsumableReference {
    part_number: String,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
struct ToolReference {
    name: String,
    setting: Option<String>,
}

// Inventory structures
#[derive(Debug, Deserialize, Clone)]
struct InventoryPart {
    part_number: String,
    description: String,
    supplier: String,
    unit_cost: f64,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryConsumable {
    part_number: String,
    description: String,
    supplier: String,
    unit_cost: f64,
}

#[derive(Debug, Deserialize, Clone)]
struct InventoryTool {
    name: String,
    brand: String,
}

#[derive(Debug, Clone)]
struct BomItem {
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
    let re = Regex::new(r"(?i)^##\s+Step\s+(\d+):?.*$").unwrap();

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
                    let parts = section_metadata.parts.as_deref().unwrap_or_default();
                    let consumables = section_metadata.consumables.as_deref().unwrap_or_default();
                    let tools = section_metadata.tools.as_deref().unwrap_or_default();

                    let parts_table = generate_parts_table(parts, inventory);
                    let consumables_table = generate_consumables_table(consumables, inventory);
                    let tools_table = generate_tools_table(tools, inventory);

                    if !parts_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(parts_table.lines().map(|s| s.to_string()));
                    }
                    if !consumables_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(consumables_table.lines().map(|s| s.to_string()));
                    }
                    if !tools_table.is_empty() {
                        result.push("".to_string()); // Empty line
                        result.extend(tools_table.lines().map(|s| s.to_string()));
                    }

                    let has_tables = !parts_table.is_empty()
                        || !consumables_table.is_empty()
                        || !tools_table.is_empty();
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

fn generate_parts_table(parts: &[PartReference], inventory: &Inventory) -> String {
    if parts.is_empty() {
        return String::new();
    }

    let mut table = String::from("<details>\n<summary><strong>🔩 Parts</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Part Number</th><th>Description</th><th>Quantity</th><th>Supplier</th></tr>\n</thead>\n<tbody>\n");

    for part_ref in parts {
        if let Some(part) = inventory.parts.get(&part_ref.part_number) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                part.part_number, part.description, part_ref.quantity, part.supplier
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Part not found in inventory</td><td>{}</td><td>-</td></tr>\n",
                part_ref.part_number, part_ref.quantity
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_consumables_table(
    consumables: &[ConsumableReference],
    inventory: &Inventory,
) -> String {
    if consumables.is_empty() {
        return String::new();
    }

    let mut table = String::from("<details>\n<summary><strong>🧪 Consumables</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Part Number</th><th>Description</th><th>Supplier</th></tr>\n</thead>\n<tbody>\n");

    for consumable_ref in consumables {
        if let Some(consumable) = inventory.consumables.get(&consumable_ref.part_number) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                consumable.part_number, consumable.description, consumable.supplier
            ));
        } else {
            table.push_str(&format!(
                "<tr><td>{}</td><td>Consumable not found in inventory</td><td>-</td></tr>\n",
                consumable_ref.part_number
            ));
        }
    }

    table.push_str("</tbody>\n</table>\n<br>\n</details>\n\n");
    table
}

fn generate_tools_table(tools: &[ToolReference], inventory: &Inventory) -> String {
    if tools.is_empty() {
        return String::new();
    }

    let mut table = String::from("<details>\n<summary><strong>🔧 Tools</strong></summary>\n<br>\n<table style=\"margin: 0;\">\n<thead>\n<tr><th>Name</th><th>Setting</th><th>Brand</th></tr>\n</thead>\n<tbody>\n");

    for tool_ref in tools {
        if let Some(tool) = inventory.tools.get(&tool_ref.name) {
            table.push_str(&format!(
                "<tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                tool.name,
                tool_ref.setting.as_deref().unwrap_or("-"),
                tool.brand
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

fn accumulate_parts(
    parts: &[PartReference],
    inventory: &Inventory,
    all_parts: &mut HashMap<String, BomItem>,
) {
    for part_ref in parts {
        if let Some(inventory_part) = inventory.parts.get(&part_ref.part_number) {
            let key = part_ref.part_number.clone();

            all_parts
                .entry(key)
                .and_modify(|item| item.total_quantity += part_ref.quantity)
                .or_insert_with(|| BomItem {
                    part_number: inventory_part.part_number.clone(),
                    description: inventory_part.description.clone(),
                    supplier: inventory_part.supplier.clone(),
                    total_quantity: part_ref.quantity,
                    unit_cost: Some(inventory_part.unit_cost),
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
        if let Some(inventory_consumable) = inventory.consumables.get(&consumable_ref.part_number) {
            let key = consumable_ref.part_number.clone();

            // For consumables, we'll just track unique items (not quantities since they're often descriptive)
            all_consumables
                .entry(key)
                .or_insert_with(|| BomConsumableItem {
                    part_number: inventory_consumable.part_number.clone(),
                    description: inventory_consumable.description.clone(),
                    supplier: inventory_consumable.supplier.clone(),
                    unit_cost: Some(inventory_consumable.unit_cost),
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
                        brand: inventory_tool.brand.clone(),
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

fn generate_bom_file(parts: &HashMap<String, BomItem>) -> Result<(), Error> {
    let mut csv_content = String::new();

    // CSV Header
    csv_content.push_str("Part Number,Description,Supplier,Quantity,Unit Cost\n");

    // Parts section
    let mut sorted_parts: Vec<_> = parts.values().collect();
    sorted_parts.sort_by(|a, b| a.description.cmp(&b.description));

    for part in sorted_parts {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\",\"{}\",{},{:.2}\n",
            part.part_number,
            part.description,
            part.supplier,
            part.total_quantity,
            part.unit_cost.unwrap_or(0.0)
        ));
    }

    // Write BOM to CSV file
    std::fs::write("output/BOM.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write BOM CSV file: {}", e)))?;

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
    csv_content.push_str("Part Number,Description,Supplier,Unit Cost\n");

    // Consumables section - only include consumables that were actually used
    let mut sorted_consumables: Vec<_> = consumables.values().collect();
    sorted_consumables.sort_by(|a, b| a.description.cmp(&b.description));

    for consumable in sorted_consumables {
        csv_content.push_str(&format!(
            "\"{}\",\"{}\",\"{}\",{:.2}\n",
            consumable.part_number,
            consumable.description,
            consumable.supplier,
            consumable.unit_cost.unwrap_or(0.0)
        ));
    }

    // Write consumables to CSV file
    std::fs::write("output/consumables.csv", csv_content)
        .map_err(|e| Error::msg(format!("Failed to write consumables CSV file: {}", e)))?;

    Ok(())
}
