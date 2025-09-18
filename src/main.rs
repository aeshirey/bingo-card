use std::collections::HashSet;

use levenshtein::levenshtein;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};

fn main() {
    let mut free_square = "FREE SQUARE".to_string();
    let mut distance_limit = 3;
    let mut people = vec!["Alice".to_string(), "Bob".to_string()];

    for arg in std::env::args().skip(1) {
        if let Some(d) = arg.strip_prefix("--dist=") {
            distance_limit = d.parse().unwrap();
        } else if let Some(pp) = arg.strip_prefix("--people=") {
            people = pp.split(',').map(|s| s.trim().to_string()).collect();
        } else if arg == "-h" {
            println!("Usage: bingo-card [--dist=N] [--people=NAME1,NAME2,...] [FREE_SQUARE_TEXT]");
            println!();
            println!("  --dist=N               Set the minimum Levenshtein distance between tiles");
            println!("  --people=NAME1,NAME2  Comma-separated list of people to generate cards for");
            println!("  FREE_SQUARE_TEXT      Text to use for the free square (default: 'FREE SQUARE')");
            return;
        } else if !arg.starts_with('-') {
            free_square = arg;
        }
    }

    let mut workbook = Workbook::new();
    let tiles = load_tiles();

    check_tiles(&tiles, distance_limit);

    for person in &people {
        generate_for_person(&mut workbook, person, &tiles, &free_square);
    }

    workbook.save("bingo.xlsx").unwrap();
}

fn check_tiles(tiles: &HashSet<String>, distance_limit: usize) {
    let tiles = tiles.iter().collect::<Vec<_>>();

    for i in 0..tiles.len() - 1 {
        for j in i + 1..tiles.len() {
            let a = tiles[i];
            let b = tiles[j];

            if a == b {
                println!("Duplicate tile: {a}");
                continue;
            }

            let dist = levenshtein(a, b);

            if dist <= distance_limit {
                println!("Similar tiles (dist={dist}):");
                println!("  1: {a}");
                println!("  2: {b}");
            }
        }
    }
}

/// Loads lines from 'tiles.txt' with one tile per line.
///
/// Lines may include a slash and 'n' that will be turned into a newline for cleaner, wrapped
/// tiles.
fn load_tiles() -> HashSet<String> {
    use std::io::BufRead as _;
    let fs = std::fs::File::open("tiles.txt").unwrap();
    let reader = std::io::BufReader::new(fs);
    reader
        .lines()
        .map(|line| line.unwrap().trim().replace("\\n", "\n"))
        .filter(|line| !line.is_empty())
        .collect()
}

fn generate_for_person(wb: &mut Workbook, name: &str, tiles: &HashSet<String>, free_square: &str) {
    const HEADER_ROWS: u32 = 2;
    const LEFT_COLS: u16 = 1;
    const CELL_DIMENSIONS: u16 = 150;

    let sheet = wb.add_worksheet();
    sheet.set_name(name).unwrap();
    sheet.set_landscape();
    sheet.set_margins(0.2, 0.2, 0.2, 0.2, 0.2, 0.2);

    let dotted_fmt = Format::new()
        .set_border(FormatBorder::Medium)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_color(Color::RGB(0x333333));

    // set the header
    let header_txt = format!("SUMO BINGO! - {name}");
    sheet
        .merge_range(1, LEFT_COLS, 1, 4 + LEFT_COLS, &header_txt, &dotted_fmt)
        .unwrap();

    for i in 0..5 {
        sheet
            .set_column_width_pixels(i + LEFT_COLS, CELL_DIMENSIONS)
            .unwrap();
        sheet
            .set_row_height_pixels(i as u32 + HEADER_ROWS, CELL_DIMENSIONS)
            .unwrap();
    }

    let mut tiles = tiles.iter().collect::<Vec<_>>();

    // randomize the order of `tiles`:
    use rand::seq::SliceRandom as _;
    let mut rng = rand::rng();
    tiles.shuffle(&mut rng);

    for (i, tile) in tiles.iter().enumerate() {
        let row = i as u32 % 5 + HEADER_ROWS;
        let col = i as u16 / 5 + LEFT_COLS;

        if row == 2 + HEADER_ROWS && col == 2 + LEFT_COLS {
            // center square
            let center_fmt = Format::new()
                .set_bold()
                .set_border(FormatBorder::Medium)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_text_wrap()
                .set_font_color(Color::White)
                .set_background_color(Color::RGB(0x222222));

            sheet
                .write_string_with_format(row, col, free_square, &center_fmt)
                .unwrap();
        } else {
            sheet
                .write_string_with_format(row, col, *tile, &dotted_fmt)
                .unwrap();
        }
    }
}
