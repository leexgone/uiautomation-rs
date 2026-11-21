# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`uiautomation-rs` is a Rust crate that wraps the Windows UI Automation API. It provides idiomatic Rust bindings for discovering, accessing, and filtering UI elements, plus keyboard/mouse input simulation, clipboard operations, process management, and UI automation events.

**Key crates:**
- `uiautomation` - main library (Rust bindings for `IUIAutomation`)
- `uiautomation_derive` - procedural macros for control pattern derive traits

## Repository Conventions

- Everything committed to this repository is written in **English**: code comments, documentation, workflow names and step labels, script messages, and **git commit messages**.
- Commit messages: short imperative subject in English, optionally with a conventional prefix seen in the history (`ci:`, `docs:`, `chore:`, `feat:`, `fix:`), e.g. `docs: document the release procedure`.
- Requirement/design/plan working documents produced during agent sessions may be written in Chinese for the maintainer's review, but they are process artifacts and are not committed to the repository.

## Build & Development Commands

```bash
# Build (default features: input, control)
cargo build -p uiautomation

# Build with all features
cargo build -p uiautomation --all-features

# Build for the docs.rs target
cargo build -p uiautomation --target x86_64-pc-windows-msvc --all-features

# Run tests (Windows only - requires an interactive desktop session)
# Tests must run single-threaded; several cases drive real UI (Notepad, focus, clipboard).
cargo test -p uiautomation -- --test-threads=1

# Run a single test
cargo test -p uiautomation -- test_element_properties -- --test-threads=1

# Run a sample
cargo run -p uia_print
cargo run -p uia_notepad

# Check without building
cargo check -p uiautomation

# Check the entire workspace (all crates, samples and test programs, all features)
cargo check --workspace --all-targets --all-features

# Release-script offline tests (no toolchain needed; also run in CI)
pwsh scripts/test-check-versions.ps1
pwsh scripts/test-crate-index.ps1

# Format
cargo fmt

# Clippy
cargo clippy -p uiautomation --all-features
```

**Important:** This crate is Windows-only. It depends on `windows` crate Win32 APIs and will not compile on Linux/macOS. The docs.rs target is `x86_64-pc-windows-msvc`; all CI jobs run on `windows-latest`.

Tests require a real interactive desktop (Session 1). Many bundled tests are demos that open Notepad, send keystrokes, or rely on the Chinese input method, and a few use hard-coded window handles; they are **not** run in CI. CI performs formatting, linting, compile checks and the release-script offline tests only.

## Architecture

### Module Structure (`crates/uiautomation/src/`)

| Module | Feature gate | Purpose |
|---|---|---|
| `core` | always | `UIAutomation`, `UIElement`, `UITreeWalker`, `UIMatcher` - the main entry points |
| `types` | always | Enums for `ControlType`, `UIProperty`, `TreeScope`, `Rect`, `Point`, etc. |
| `errors` | always | `Error` and `Result` types |
| `variants` | always | `Variant` wrapper for COM `VARIANT` types, `SafeArray` |
| `filters` | always | `MatcherFilter` trait and built-in filters (`NameFilter`, `ClassNameFilter`, `ControlTypeFilter`) |
| `logs` | always | `pub(crate)` internal `log_debug!` / `log_error!` macros |
| `patterns` | `pattern` | Wrappers for all 30+ UIA control pattern interfaces (Invoke, Value, Window, Text, etc.) |
| `controls` | `control` | Typed control wrappers (`WindowControl`, `EditControl`, etc.) using derive macros |
| `actions` | `control` | Action traits implemented by controls |
| `inputs` | `input` | `Keyboard` and `Mouse` simulation |
| `processes` | `process` | `Process` struct for spawn/wait |
| `dialogs` | `dialog` | Message box utilities |
| `events` | `event` | UIA event handlers (focus changed, property changed, structure changed) |
| `clipboards` | `clipboard` | Clipboard snapshot/restore |
| `screenshots` | `screenshot` | Screen/region capture to bitmaps and PNG (uses `png` + GDI) |

### Key Abstractions

- **`UIAutomation`** - entry point, wraps `IUIAutomation`. Handles COM initialization. Creates walkers, matchers, conditions.
- **`UIElement`** - wraps `IUIAutomationElement`. Provides property getters (name, classname, bounding rect, etc.), `send_keys()`, `click()`, pattern access via `get_pattern::<T>()`.
- **`UITreeWalker`** - traverses the UI Automation tree (parent/child/sibling navigation).
- **`UIMatcher`** - builder-pattern search with filters (`name()`, `classname()`, `control_type()`, `filter_fn()`), timeout-based retries, and depth-limited tree search.
- **`UICondition`** family - wraps COM condition objects for `FindFirst`/`FindAll`.

### Derive Macros (`crates/uiautomation_derive/`)

- `EnumConvert` - generates `TryFrom<i32>` / `Into<i32>` for typed enums backed by integer constants.
- `pattern_as` - attribute macro that generates `UIPattern` impl and pattern type mapping.
- Control derive macros (`Invoke`, `Window`, `Value`, `Toggle`, `Selection`, `Scroll`, `Text`, etc.) - each generates methods for a specific UIA control pattern on the target struct.

### Feature Dependencies

```
default = ["input", "control"]
control -> pattern -> uiautomation_derive/{action_derive, control_derive, pattern_derive}
```

When adding a new feature, ensure the required `windows` sub-features are declared in `Cargo.toml`.

## Adding a New Control Pattern

1. Define the wrapper struct in [patterns.rs](crates/uiautomation/src/patterns.rs) with `#[pattern_as(UIPatternType::X, IUIAutomationYPattern)]`.
2. Add the corresponding derive macro in [uiautomation_derive/src/action_derives.rs](crates/uiautomation_derive/src/action_derives.rs) if the pattern has actions.
3. Register the derive macro in [uiautomation_derive/src/lib.rs](crates/uiautomation_derive/src/lib.rs).
4. Add a typed control wrapper in [controls.rs](crates/uiautomation/src/controls.rs) using `#[derive(YourPattern)]`.
5. Gate everything behind the appropriate feature flag.

## Important Notes

- All COM interop uses the `windows` crate (not `winapi`). Check the `windows` workspace dependency for the current version.
- `UIElement` has both "current" and "cached" property getters - cached versions require `UICacheRequest` and are faster for bulk access.
- The `UIMatcher` retries finds with a configurable `timeout` (default 3000ms) and `interval` (default 100ms). Its timeout is measured with a monotonic clock (`std::time::Instant`), so changing the system clock during a search does not affect it.
- Edition is 2024 - use the latest Rust idioms.

## CI

GitHub Actions workflows live in `.github/workflows/` and run on `windows-latest` only:

- **`ci.yml`** - runs on pushes to `main` and on pull requests targeting `main`:
  1. release-script offline tests;
  2. version dependency consistency check (`scripts/check-versions.ps1`);
  3. `cargo fmt --all --check`;
  4. `cargo clippy --workspace --all-targets --all-features -- -D warnings`;
  5. `cargo check --workspace --all-targets --all-features`.
- **`release.yml`** - runs only on `v*` tags; see Release below.

Release-script helper files in `scripts/`:

| File | Purpose |
|---|---|
| `check-versions.ps1` | Enforces the exact derive pin and, with `-ExpectedTag`, the tag/version match |
| `CrateIndex.psm1` | crates.io sparse-index helpers (URL rule, NDJSON parsing, 404 handling) |
| `publish-crate.ps1` | Idempotent publish of one crate (skip if already published, `--dry-run`, wait for index) |
| `test-check-versions.ps1`, `test-crate-index.ps1` | Offline tests for the two above (run in CI) |

## Release

The two crates are versioned independently, but **every `uiautomation_derive` release is accompanied by a `uiautomation` release** - the main crate pins the derive crate to an **exact** version. A single repository tag drives one release of both crates; the tag version follows the main crate (`v<uiautomation version>`).

Before tagging:

1. Bump `version` in `crates/uiautomation_derive/Cargo.toml` (if derive changed) and/or `crates/uiautomation/Cargo.toml`.
2. In `crates/uiautomation/Cargo.toml`, update the derive dependency to the exact new version, e.g.:
   ```toml
   uiautomation_derive = { version = "=0.7.10", path = "../uiautomation_derive" }
   ```
   Only `=X.Y.Z` is accepted; `^X.Y.Z` or a bare `X.Y.Z` fails CI.
3. Merge to `main` and wait for CI to pass.

Then tag and push:

```bash
git tag v<uiautomation-version>     # e.g. v0.25.2
git push origin v<uiautomation-version>
```

`release.yml` then: re-runs the offline tests and version/tag checks -> `cargo check --workspace --all-targets --all-features` -> publishes `uiautomation_derive` first, waits for it to appear in the crates.io index -> publishes `uiautomation` -> creates a GitHub Release with auto-generated notes. Already-published versions are skipped, so re-running a tag is safe.

Prerequisite: a crates.io publish-scoped token stored as the repository secret `CARGO_REGISTRY_TOKEN` (Settings -> Secrets and variables -> Actions).

## Known Tooling Debt

`cargo fmt --all --check` and `cargo clippy --workspace --all-targets --all-features -- -D warnings` are now enforced in CI, and the whole tree is formatted and clippy-clean. Keep both gates green: run `cargo fmt` before committing, and prefer `cargo clippy --fix` for new lints.

Remaining item: the interactive UI tests still cannot run in CI (see the note above). Adding headless smoke tests and marking the desktop-dependent cases `#[ignore]` is tracked separately (see `docs/superpowers/specs/2026-11-20-ci-release-pipeline-design.md`, section 6).
