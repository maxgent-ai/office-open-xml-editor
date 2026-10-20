use std::ffi::OsString;

use anyhow::Result;
use rmcp::{transport::stdio, ServiceExt};
use tracing_subscriber::{self, EnvFilter};

mod parsed_cache;
mod path_scope;
mod server;
mod tools;

const USAGE: &str =
    "usage: ooxml-mcp-server [--allowed-root <DIR>]...\n       ooxml-mcp-server --version";

#[derive(Debug, PartialEq, Eq)]
enum Command {
    Version,
    Serve { allowed_roots: Vec<OsString> },
}

/// Parse the command line. Unknown arguments are rejected rather than
/// ignored, so a misspelt scope flag cannot silently leave the scope off.
fn parse_args(args: impl IntoIterator<Item = OsString>) -> Result<Command, String> {
    let args: Vec<OsString> = args.into_iter().skip(1).collect();
    if let [flag] = args.as_slice() {
        if flag == "--version" || flag == "-V" {
            return Ok(Command::Version);
        }
    }
    let missing_root = || format!("--allowed-root requires a directory\n{USAGE}");
    let mut allowed_roots = Vec::new();
    let mut args = args.into_iter();
    while let Some(arg) = args.next() {
        if arg == "--allowed-root" {
            match args.next() {
                Some(root) if !root.is_empty() => allowed_roots.push(root),
                _ => return Err(missing_root()),
            }
            continue;
        }
        match arg
            .to_str()
            .and_then(|arg| arg.strip_prefix("--allowed-root="))
        {
            Some("") => return Err(missing_root()),
            Some(root) => allowed_roots.push(root.into()),
            None => {
                return Err(format!(
                    "unrecognized argument '{}'\n{USAGE}",
                    arg.to_string_lossy()
                ))
            }
        }
    }
    Ok(Command::Serve { allowed_roots })
}

#[tokio::main]
async fn main() -> Result<()> {
    let allowed_roots = match parse_args(std::env::args_os()) {
        Ok(Command::Version) => {
            println!("ooxml-mcp-server {}", env!("CARGO_PKG_VERSION"));
            return Ok(());
        }
        Ok(Command::Serve { allowed_roots }) => allowed_roots,
        Err(message) => {
            eprintln!("ooxml-mcp-server: {message}");
            std::process::exit(2);
        }
    };

    tracing_subscriber::fmt()
        .with_env_filter(EnvFilter::from_default_env())
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    tracing::info!("Starting ooxml-mcp-server");

    let env_roots = std::env::var_os(path_scope::ALLOWED_ROOTS_ENV);
    let scope =
        path_scope::from_config(allowed_roots, env_roots.as_deref()).map(|(scope, warnings)| {
            for warning in warnings {
                tracing::warn!("{warning}");
            }
            scope
        });
    if scope.is_none() {
        tracing::info!("No allowed roots configured; path-based tools may read any local file");
    }
    path_scope::install(scope);

    let service = server::OoxmlServer::new()
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("Server error: {:?}", e))?;

    service.waiting().await?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{parse_args, Command};
    use std::ffi::OsString;

    fn parse(args: &[&str]) -> Result<Command, String> {
        parse_args(
            std::iter::once("server")
                .chain(args.iter().copied())
                .map(OsString::from),
        )
    }

    #[test]
    fn recognizes_only_a_standalone_version_flag() {
        assert_eq!(parse(&["--version"]), Ok(Command::Version));
        assert_eq!(parse(&["-V"]), Ok(Command::Version));
        assert_eq!(
            parse(&[]),
            Ok(Command::Serve {
                allowed_roots: Vec::new()
            })
        );
        assert!(parse(&["--version", "extra"]).is_err());
    }

    #[test]
    fn collects_repeated_allowed_roots_and_rejects_unknown_or_empty_flags() {
        assert_eq!(
            parse(&["--allowed-root", "/a", "--allowed-root=/b"]),
            Ok(Command::Serve {
                allowed_roots: vec!["/a".into(), "/b".into()]
            })
        );
        assert!(parse(&["--allowed-roots", "/a"]).is_err());
        assert!(parse(&["--allowed-root"]).is_err());
        assert!(parse(&["--allowed-root="]).is_err());
    }
}
