//! Optional allowed-roots scope for every path-based tool.
//!
//! Tool arguments come from an AI agent, and an agent can be steered by the
//! content of the documents it reads. Without a scope the server opens any
//! readable regular file it is given. When roots are configured (the
//! `--allowed-root` CLI flag or the `OOXML_MCP_ALLOWED_ROOTS` environment
//! variable), a requested path is canonicalized first — resolving `..`
//! components and every symlink — and must then lie inside one of the
//! canonicalized roots. Anything else is rejected before the file is opened.
//!
//! The rejection message is identical for a path outside the roots and for a
//! path that cannot be resolved at all, so the scope never reveals whether a
//! file exists outside it.
//!
//! On Windows, resolving a path is itself file-system access: canonicalizing
//! `\\server\share\x.docx` opens an SMB connection (and may send NTLM
//! credentials) before any containment check could run. So when roots are
//! configured, UNC paths (`\\server\share`, `//server/share`,
//! `\\?\UNC\...`) and device or object-manager namespaces (`\\.\`,
//! `\\?\` other than a drive, `\??\`) are rejected lexically before any
//! file-system call. A UNC request is admitted to the canonical check only
//! when it lexically lies inside a configured root on that same share.
//!
//! Limitation: the check and the later open are separate system calls. The
//! scope stops an agent from being steered to other files; it is not a
//! defence against a concurrent local process that can rewrite directories
//! inside a root between the check and the open. Likewise, a symlink or
//! junction planted inside a root that points at a network share is followed
//! by canonicalization; that requires write access inside the root.
//!
//! Without configuration the scope is off, preserving the behaviour of
//! standalone installations that rely on reading any local path.

use std::ffi::{OsStr, OsString};
use std::fmt;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

/// Environment variable holding allowed roots, separated like `PATH`
/// (`:` on Unix, `;` on Windows).
pub const ALLOWED_ROOTS_ENV: &str = "OOXML_MCP_ALLOWED_ROOTS";

/// Canonicalized allowed roots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PathScope {
    roots: Vec<PathBuf>,
    /// Lexical `[server, share, dirs...]` keys of roots that live on a UNC
    /// share, lowercased (Windows compares these names case-insensitively).
    network_roots: Vec<Vec<String>>,
    /// Apply the Windows lexical gate. `cfg!(windows)` outside tests.
    windows: bool,
}

/// A request rejected by a configured [`PathScope`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OutsideAllowedRoots {
    pub path: String,
}

impl fmt::Display for OutsideAllowedRoots {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "Cannot read '{}': path_outside_allowed_roots: the path does not exist or is outside the allowed roots configured for this server",
            self.path
        )
    }
}

impl std::error::Error for OutsideAllowedRoots {}

impl PathScope {
    /// Canonicalize each root. A root that cannot be resolved is skipped and
    /// reported in the returned warnings; the scope stays in force, so a
    /// scope whose every root is invalid denies every request.
    pub fn new<I, P>(roots: I) -> (Self, Vec<String>)
    where
        I: IntoIterator<Item = P>,
        P: AsRef<Path>,
    {
        let mut canonical = Vec::new();
        let mut network_roots = Vec::new();
        let mut warnings = Vec::new();
        for root in roots {
            let root = root.as_ref();
            // A relative root would silently depend on the launcher's working
            // directory, so only absolute roots are accepted.
            if !root.is_absolute() {
                warnings.push(format!(
                    "ignoring allowed root '{}': not an absolute path",
                    root.display()
                ));
                continue;
            }
            match fs::canonicalize(root) {
                Ok(path) => {
                    // Record both spellings of a share root, e.g.
                    // `\\server\share\dir` and `\\?\UNC\server\share\dir`.
                    for spelling in [root, path.as_path()] {
                        if let Some(key) = network_components(&spelling.to_string_lossy()) {
                            if !network_roots.contains(&key) {
                                network_roots.push(key);
                            }
                        }
                    }
                    if !canonical.contains(&path) {
                        canonical.push(path);
                    }
                }
                Err(error) => warnings.push(format!(
                    "ignoring allowed root '{}': {}",
                    root.display(),
                    error
                )),
            }
        }
        (
            Self {
                roots: canonical,
                network_roots,
                windows: cfg!(windows),
            },
            warnings,
        )
    }

    /// Lexical admission that must pass before any file-system call. On
    /// Windows it keeps canonicalization from reaching a network share or a
    /// device namespace outside the configured roots.
    fn lexically_admissible(&self, requested: &str) -> bool {
        if !self.windows {
            return true;
        }
        match windows_prefix(requested) {
            WindowsPrefix::Local => true,
            WindowsPrefix::Device => false,
            WindowsPrefix::Network => network_components(requested).is_some_and(|request| {
                self.network_roots
                    .iter()
                    .any(|root| request.starts_with(root))
            }),
        }
    }

    /// Resolve `requested` to its canonical path when it lies inside a root.
    pub fn resolve(&self, requested: &str) -> Result<PathBuf, OutsideAllowedRoots> {
        let denied = || OutsideAllowedRoots {
            path: requested.to_string(),
        };
        if !self.lexically_admissible(requested) {
            return Err(denied());
        }
        let canonical = fs::canonicalize(requested).map_err(|_| denied())?;
        // `Path::starts_with` compares whole components, so a sibling such as
        // `/data/project-other` is not inside the root `/data/project`.
        if self.roots.iter().any(|root| canonical.starts_with(root)) {
            Ok(canonical)
        } else {
            Err(denied())
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WindowsPrefix {
    /// A drive or relative path, including the verbatim drive form `\\?\C:\`.
    Local,
    /// A UNC path: `\\server\share`, `//server/share`, `\\?\UNC\server\share`.
    Network,
    /// `\\.\...`, any other `\\?\...` (GLOBALROOT, volume GUIDs) or the
    /// NT object-manager prefix `\??\`.
    Device,
}

fn is_separator(c: char) -> bool {
    c == '\\' || c == '/'
}

/// Classify `path` by its Windows prefix, purely lexically (no file-system
/// access), following the Win32 file path forms (UNC, `\\?\` verbatim,
/// `\\.\` device) that Windows path normalization recognizes.
fn windows_prefix(path: &str) -> WindowsPrefix {
    let chars: Vec<char> = path.chars().take(4).collect();
    let at = |index: usize| chars.get(index).copied();
    let sep = |index: usize| at(index).is_some_and(is_separator);
    if sep(0) && at(1) == Some('?') && at(2) == Some('?') && sep(3) {
        return WindowsPrefix::Device;
    }
    if !(sep(0) && sep(1)) {
        return WindowsPrefix::Local;
    }
    match at(2) {
        Some(kind @ ('?' | '.')) if sep(3) => {
            // The first four characters are ASCII, so byte 4 is a boundary.
            let tail = &path[4..];
            let mut tail_chars = tail.chars();
            let drive = tail_chars.next().is_some_and(|c| c.is_ascii_alphabetic())
                && tail_chars.next() == Some(':')
                && tail_chars.next().is_none_or(is_separator);
            let unc = tail.get(..3).is_some_and(|p| p.eq_ignore_ascii_case("UNC"))
                && tail
                    .get(3..)
                    .is_some_and(|rest| rest.chars().next().is_none_or(is_separator));
            if kind == '?' && drive {
                WindowsPrefix::Local
            } else if kind == '?' && unc {
                WindowsPrefix::Network
            } else {
                WindowsPrefix::Device
            }
        }
        _ => WindowsPrefix::Network,
    }
}

/// Lexical `[server, share, components...]` of a UNC path, lowercased, with
/// `.` dropped and `..` applied. `None` for a non-UNC path, a path without a
/// server and share, or `..` that would climb out of the share.
fn network_components(path: &str) -> Option<Vec<String>> {
    if windows_prefix(path) != WindowsPrefix::Network {
        return None;
    }
    // `\\?\UNC` is seven ASCII bytes; otherwise skip the two separators.
    let rest = if path[2..].starts_with('?') {
        &path[7..]
    } else {
        &path[2..]
    };
    let mut components: Vec<String> = Vec::new();
    for component in rest.split(is_separator) {
        match component {
            "" | "." => {}
            ".." => {
                if components.len() <= 2 {
                    return None;
                }
                components.pop();
            }
            name => components.push(name.to_lowercase()),
        }
    }
    (components.len() >= 2).then_some(components)
}

static SCOPE: OnceLock<Option<PathScope>> = OnceLock::new();

/// Install the process-wide scope. Called once from `main` before serving.
pub fn install(scope: Option<PathScope>) {
    let _ = SCOPE.set(scope);
}

/// Canonicalize a tool's path argument, enforcing the configured scope.
pub fn canonicalize(requested: &str) -> Result<PathBuf, String> {
    match SCOPE.get_or_init(|| None) {
        Some(scope) => scope.resolve(requested).map_err(|error| error.to_string()),
        None => fs::canonicalize(requested)
            .map_err(|error| format!("Cannot read '{}': {}", requested, error)),
    }
}

/// Allowed roots requested by the environment variable, if it is set and not
/// empty. Empty entries (for example a trailing separator) are ignored.
pub fn roots_from_env(value: Option<&OsStr>) -> Option<Vec<PathBuf>> {
    let value = value.filter(|value| !value.is_empty())?;
    Some(
        std::env::split_paths(value)
            .filter(|path| !path.as_os_str().is_empty())
            .collect(),
    )
}

/// Build the scope from CLI roots plus the environment variable. Returns
/// `None` when neither supplies a root, which leaves the scope off.
pub fn from_config(
    cli_roots: Vec<OsString>,
    env_value: Option<&OsStr>,
) -> Option<(PathScope, Vec<String>)> {
    let env_roots = roots_from_env(env_value);
    if cli_roots.is_empty() && env_roots.is_none() {
        return None;
    }
    let roots = cli_roots
        .into_iter()
        .map(PathBuf::from)
        .chain(env_roots.unwrap_or_default());
    Some(PathScope::new(roots))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::SystemTime;

    struct TempTree(PathBuf);

    impl TempTree {
        fn new(tag: &str) -> Self {
            let root = std::env::temp_dir().join(format!(
                "ooxml-mcp-scope-{tag}-{}-{}",
                std::process::id(),
                SystemTime::now()
                    .duration_since(SystemTime::UNIX_EPOCH)
                    .unwrap()
                    .as_nanos()
            ));
            fs::create_dir_all(root.join("allowed/nested")).unwrap();
            fs::create_dir_all(root.join("allowed-sibling")).unwrap();
            fs::create_dir_all(root.join("outside")).unwrap();
            fs::write(root.join("allowed/nested/inside.docx"), b"inside").unwrap();
            fs::write(root.join("allowed-sibling/sibling.docx"), b"sibling").unwrap();
            fs::write(root.join("outside/secret.docx"), b"secret").unwrap();
            Self(root)
        }

        fn path(&self, relative: &str) -> String {
            self.0.join(relative).to_str().unwrap().to_string()
        }
    }

    impl Drop for TempTree {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    fn scope(tree: &TempTree) -> PathScope {
        let (scope, warnings) = PathScope::new([tree.0.join("allowed")]);
        assert!(warnings.is_empty());
        scope
    }

    #[test]
    fn accepts_a_file_inside_a_root_and_returns_its_canonical_path() {
        let tree = TempTree::new("inside");
        let resolved = scope(&tree)
            .resolve(&tree.path("allowed/nested/inside.docx"))
            .unwrap();
        assert_eq!(
            resolved,
            fs::canonicalize(tree.0.join("allowed/nested/inside.docx")).unwrap()
        );
        // `..` that stays inside the root is fine.
        assert!(scope(&tree)
            .resolve(&tree.path("allowed/nested/../nested/inside.docx"))
            .is_ok());
    }

    #[test]
    fn rejects_outside_paths_traversal_and_component_prefix_siblings() {
        let tree = TempTree::new("outside");
        let scope = scope(&tree);
        for requested in [
            tree.path("outside/secret.docx"),
            tree.path("allowed/../outside/secret.docx"),
            tree.path("allowed/nested/../../outside/secret.docx"),
            tree.path("allowed-sibling/sibling.docx"),
        ] {
            let error = scope.resolve(&requested).unwrap_err();
            assert_eq!(error.path, requested);
            assert!(error.to_string().contains("path_outside_allowed_roots"));
        }
    }

    #[test]
    fn denial_does_not_reveal_whether_an_outside_path_exists() {
        let tree = TempTree::new("oracle");
        let scope = scope(&tree);
        let existing = scope
            .resolve(&tree.path("outside/secret.docx"))
            .unwrap_err()
            .to_string();
        let missing = scope
            .resolve(&tree.path("outside/missing.docx"))
            .unwrap_err()
            .to_string();
        assert_eq!(
            existing.replace("secret", "X"),
            missing.replace("missing", "X")
        );
    }

    #[cfg(unix)]
    #[test]
    fn rejects_symlinks_that_escape_the_root_and_accepts_ones_that_stay_inside() {
        use std::os::unix::fs::symlink;
        let tree = TempTree::new("symlink");
        symlink(
            tree.0.join("outside/secret.docx"),
            tree.0.join("allowed/link.docx"),
        )
        .unwrap();
        symlink(tree.0.join("outside"), tree.0.join("allowed/dir-link")).unwrap();
        symlink(
            tree.0.join("allowed/nested/inside.docx"),
            tree.0.join("allowed/inner-link.docx"),
        )
        .unwrap();
        let scope = scope(&tree);
        assert!(scope.resolve(&tree.path("allowed/link.docx")).is_err());
        assert!(scope
            .resolve(&tree.path("allowed/dir-link/secret.docx"))
            .is_err());
        assert_eq!(
            scope
                .resolve(&tree.path("allowed/inner-link.docx"))
                .unwrap(),
            fs::canonicalize(tree.0.join("allowed/nested/inside.docx")).unwrap()
        );
    }

    #[test]
    fn classifies_windows_network_and_device_paths_lexically() {
        use WindowsPrefix::*;
        for (path, expected) in [
            (r"C:\docs\a.docx", Local),
            (r"docs\a.docx", Local),
            (r"\docs\a.docx", Local),
            (r"\\?\C:\docs\a.docx", Local),
            (r"\\?\c:", Local),
            (r"\\attacker\share\x.docx", Network),
            ("//attacker/share/x.docx", Network),
            (r"\/attacker\share", Network),
            (r"\\?\UNC\attacker\share\x.docx", Network),
            (r"\\?\unc\attacker\share", Network),
            (r"\\.\pipe\x", Device),
            (r"\\.\C:\x.docx", Device),
            (r"\\?\GLOBALROOT\Device\Mup\attacker\share", Device),
            (
                r"\\?\Volume{00000000-0000-0000-0000-000000000000}\x",
                Device,
            ),
            (r"\\?\UNCX\attacker", Device),
            (r"\\?\UNé", Device),
            (r"\??\UNC\attacker\share", Device),
            ("/??/C:/x", Device),
        ] {
            assert_eq!(windows_prefix(path), expected, "{path}");
        }
        assert_eq!(
            network_components(r"\\?\UNC\Srv\Share\.\Dir\..\Doc.docx"),
            Some(vec!["srv".into(), "share".into(), "doc.docx".into()])
        );
        assert_eq!(network_components(r"\\srv\share\..\other"), None);
        assert_eq!(network_components(r"\\srv"), None);
        assert_eq!(network_components(r"C:\x"), None);
    }

    #[test]
    fn windows_gate_rejects_network_and_device_paths_before_any_file_access() {
        let windowsish = |network_roots: Vec<Vec<String>>| PathScope {
            roots: Vec::new(),
            network_roots,
            windows: true,
        };
        let local_only = windowsish(Vec::new());
        for path in [
            r"\\attacker\share\x.docx",
            "//attacker/share/x.docx",
            r"\\?\UNC\attacker\share\x.docx",
            r"\\.\pipe\x",
            r"\\?\GLOBALROOT\Device\Mup\attacker\share\x.docx",
            r"\??\UNC\attacker\share\x.docx",
        ] {
            assert!(!local_only.lexically_admissible(path), "{path}");
        }
        assert!(local_only.lexically_admissible(r"C:\work\a.docx"));
        assert!(local_only.lexically_admissible(r"\\?\C:\work\a.docx"));

        let share = windowsish(vec![network_components(r"\\srv\share\team").unwrap()]);
        for path in [
            r"\\SRV\Share\team\a.docx",
            "//srv/share/team/sub/a.docx",
            r"\\?\UNC\srv\share\team\a.docx",
        ] {
            assert!(share.lexically_admissible(path), "{path}");
        }
        for path in [
            r"\\srv\share\other\a.docx",
            r"\\srv\share\team\..\other\a.docx",
            r"\\srv\share\teammate\a.docx",
            r"\\attacker\share\team\a.docx",
            r"\\srv\share",
        ] {
            assert!(!share.lexically_admissible(path), "{path}");
        }
        // Off Windows these are ordinary local names.
        let posix = PathScope {
            windows: false,
            ..local_only
        };
        assert!(posix.lexically_admissible("//attacker/share/x.docx"));
    }

    #[cfg(windows)]
    #[test]
    fn windows_resolve_denies_unc_requests_outside_the_roots() {
        let tree = TempTree::new("unc");
        let scope = scope(&tree);
        for path in [
            r"\\attacker.invalid\share\x.docx",
            r"\\?\UNC\attacker.invalid\share\x.docx",
            r"\\.\pipe\ooxml",
        ] {
            assert!(scope
                .resolve(path)
                .unwrap_err()
                .to_string()
                .contains("path_outside_allowed_roots"));
        }
    }

    #[test]
    fn configuration_is_opt_in_and_invalid_roots_fail_closed() {
        assert_eq!(from_config(Vec::new(), None), None);
        assert_eq!(from_config(Vec::new(), Some(OsStr::new(""))), None);

        let tree = TempTree::new("config");
        let env = std::env::join_paths([tree.0.join("allowed"), tree.0.join("missing")]).unwrap();
        let (scope, warnings) = from_config(Vec::new(), Some(&env)).unwrap();
        assert_eq!(warnings.len(), 1);
        assert!(scope
            .resolve(&tree.path("allowed/nested/inside.docx"))
            .is_ok());
        assert!(scope.resolve(&tree.path("outside/secret.docx")).is_err());

        let (scope, warnings) = from_config(
            vec![tree.0.join("missing").into_os_string(), "allowed".into()],
            None,
        )
        .unwrap();
        assert_eq!(warnings.len(), 2);
        assert!(scope
            .resolve(&tree.path("allowed/nested/inside.docx"))
            .is_err());
    }
}
