use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, ValueEnum};

mod commands;
mod render;

/// Fast, agent-friendly Excel previews.
///
/// `wolfxl peek <file>` prints a styled, token-efficient view of a workbook —
/// box / text / csv / json output, sheet selection, row and width caps.
/// `wolfxl map <file>` prints a one-page summary of every sheet.
#[derive(Parser, Debug)]
#[command(name = "wolfxl", version, about, long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(clap::Subcommand, Debug)]
enum Command {
    /// Print a preview of a spreadsheet.
    Peek(PeekArgs),
    /// Print a one-page workbook overview (sheets, dims, headers, named ranges).
    Map(MapArgs),
}

#[derive(clap::Args, Debug)]
struct MapArgs {
    /// Path to the workbook (.xlsx).
    file: PathBuf,

    /// Output format.
    #[arg(short = 'f', long = "format", default_value = "json")]
    format: MapFormat,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub enum MapFormat {
    Json,
    Text,
}

#[derive(clap::Args, Debug)]
struct PeekArgs {
    /// Path to the workbook (.xlsx).
    file: PathBuf,

    /// Sheet name (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// Maximum number of rows to display (0 = all).
    #[arg(short = 'n', long = "max-rows", default_value_t = 50)]
    max_rows: usize,

    /// Maximum column width in characters.
    #[arg(short = 'w', long = "max-width", default_value_t = 30)]
    max_width: usize,

    /// Export format. Omit for the boxed terminal preview.
    #[arg(short = 'e', long = "export")]
    export: Option<ExportFormat>,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
enum ExportFormat {
    Csv,
    Json,
    Text,
    Box,
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    let result = match cli.command {
        Command::Peek(args) => commands::peek::run(args),
        Command::Map(args) => commands::map::run(args.file, args.format),
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("error: {e:#}");
            ExitCode::from(1)
        }
    }
}
