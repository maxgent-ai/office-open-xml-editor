use anyhow::Result;
use rmcp::{transport::stdio, ServiceExt};
use tracing_subscriber::{self, EnvFilter};

mod parsed_cache;
mod server;
mod tools;

fn version_requested(args: impl IntoIterator<Item = String>) -> bool {
    matches!(
        args.into_iter().collect::<Vec<_>>().as_slice(),
        [_, flag] if flag == "--version" || flag == "-V"
    )
}

#[tokio::main]
async fn main() -> Result<()> {
    if version_requested(std::env::args()) {
        println!("ooxml-mcp-server {}", env!("CARGO_PKG_VERSION"));
        return Ok(());
    }

    tracing_subscriber::fmt()
        .with_env_filter(EnvFilter::from_default_env())
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    tracing::info!("Starting ooxml-mcp-server");

    let service = server::OoxmlServer::new()
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("Server error: {:?}", e))?;

    service.waiting().await?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::version_requested;

    #[test]
    fn recognizes_only_a_standalone_version_flag() {
        assert!(version_requested(["server".into(), "--version".into()]));
        assert!(version_requested(["server".into(), "-V".into()]));
        assert!(!version_requested(["server".into()]));
        assert!(!version_requested([
            "server".into(),
            "--version".into(),
            "extra".into(),
        ]));
    }
}
