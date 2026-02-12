use anyhow::Result;
use clap::Parser;
use spreadsheet_mcp::cli;

#[tokio::main(flavor = "current_thread")]
async fn main() -> Result<()> {
    let cli_args = cli::Cli::parse();
    cli::errors::ensure_output_supported(cli_args.format)?;
    let payload = cli::run_command(cli_args.command).await?;
    cli::output::emit_value(&payload, cli_args.format, cli_args.compact, cli_args.quiet)?;
    Ok(())
}
