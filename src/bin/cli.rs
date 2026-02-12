use anyhow::Result;
use clap::Parser;
use spreadsheet_mcp::cli;

#[tokio::main(flavor = "current_thread")]
async fn main() -> Result<()> {
    let cli_args = cli::Cli::parse();
    if let Err(error) = cli::errors::ensure_output_supported(cli_args.format) {
        emit_error_and_exit(error);
    }

    match cli::run_command(cli_args.command).await {
        Ok(payload) => {
            if let Err(error) =
                cli::output::emit_value(&payload, cli_args.format, cli_args.compact, cli_args.quiet)
            {
                emit_error_and_exit(error);
            }
            Ok(())
        }
        Err(error) => {
            emit_error_and_exit(error);
        }
    }
}

fn emit_error_and_exit(error: anyhow::Error) -> ! {
    let envelope = cli::errors::envelope_for(&error);
    let stderr = std::io::stderr();
    let mut handle = stderr.lock();
    if serde_json::to_writer(&mut handle, &envelope).is_err() {
        eprintln!("{{\"code\":\"COMMAND_FAILED\",\"message\":\"{}\"}}", error);
    } else {
        use std::io::Write;
        let _ = handle.write_all(b"\n");
    }
    std::process::exit(1)
}
