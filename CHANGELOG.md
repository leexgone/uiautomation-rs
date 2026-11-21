# Changelog

All notable changes to this project are documented in this file.

The format is inspired by [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html). Only notable changes are listed;
dependency-only maintenance releases may be omitted. Dates are the crates.io publish dates.

`uiautomation_derive` is released together with `uiautomation` and is pinned to an exact version,
so changes to the derive crate are recorded under the corresponding `uiautomation` release.

## Unreleased

## 0.25.3 - 2026-11-21

### Changed

- Formatted the whole workspace with `rustfmt` and fixed all `clippy` lints; CI now enforces
  `cargo fmt --check` and `cargo clippy -- -D warnings`.
- Replaced `impl Into<...>` with `impl From<...>`; `.into()` call sites remain compatible.
- `SafeArray::from_vector()` and `SafeArray::from_string_vector()` now accept slices (`&[T]`).
- Moved the test crates' `uiautomation` and `windows` dependencies to `[dev-dependencies]`.
- Pinned `uiautomation_derive` to the exact version `=0.7.11`.

## 0.25.2 - 2026-11-20

### Changed

- Updated `log` to `0.4.35` and `syn` to `3.0.10` (used by `uiautomation_derive`).
- Pinned `uiautomation_derive` to the exact version `=0.7.10`.

## 0.25.1 - 2026-09-04

### Changed

- Measured `UIMatcher` timeouts with a monotonic clock, so changing the system clock no longer
  affects a running search. ([#127](https://github.com/leexgone/uiautomation-rs/pull/127))

## 0.25.0 - 2026-05-05

### Added

- Added the `screenshot` feature for capturing screens and regions as bitmaps and PNG files.

## 0.24.4 - 2026-03-02

### Changed

- Updated dependencies.

## 0.24.3 - 2026-01-12

### Changed

- Updated dependencies.

## 0.24.2 - 2025-11-19

### Changed

- Updated dependencies (Cargo 1.19.1).

## 0.24.1 - 2025-10-07

### Changed

- Updated to `windows v0.62.2`.

## 0.24.0 - 2025-09-20

### Added

- Added the repetition-count syntax for special keys in input strings.
  ([#124](https://github.com/leexgone/uiautomation-rs/issues/124))

## 0.23.1 - 2025-09-05

### Changed

- Updated to `windows v0.62`.

## 0.23.0 - 2025-08-29

### Added

- Added mouse drag support. ([#123](https://github.com/leexgone/uiautomation-rs/issues/123))

## 0.22.2 - 2025-08-02

### Added

- Added `Window::get_window_visual_state()` and `Window::set_window_visual_state()`.
  ([#119](https://github.com/leexgone/uiautomation-rs/issues/119))

## 0.22.1 - 2025-07-30

### Added

- Added `UITreeWalker::get_children()` and `UITreeWalker::get_children_build_cache()`.
  ([#118](https://github.com/leexgone/uiautomation-rs/issues/118))

## 0.22.0 - 2025-07-12

### Fixed

- Fixed clipboard crashes after repeated use.
  ([#114](https://github.com/leexgone/uiautomation-rs/issues/114))

## 0.21.6 - 2025-07-07

### Added

- Added unicode input creation.
  ([#113](https://github.com/leexgone/uiautomation-rs/pull/113), thanks
  [@elsnerpaul](https://github.com/elsnerpaul))

## 0.21.5 - 2025-07-05

### Fixed

- Fixed clipboard crashes. ([#111](https://github.com/leexgone/uiautomation-rs/issues/111))

## 0.21.4 - 2025-07-04

### Changed

- No longer focuses an element before clicking it.
  ([#109](https://github.com/leexgone/uiautomation-rs/issues/109))

## 0.21.3 - 2025-06-25

### Changed

- Updated to `windows v0.61.3`.

## 0.21.2 - 2025-06-12

### Changed

- Changed the documentation settings.

## 0.21.1 - 2025-06-12

### Changed

- Updated to `windows v0.61.2`.

## 0.21.0 - 2025-06-10

### Added

- Added the `clipboard` feature, including Windows clipboard support.

## 0.20.0 - 2025-06-01

### Changed

- Refactored `Process`.

## 0.19.2 - 2025-05-23

### Added

- Added more key names. ([#106](https://github.com/leexgone/uiautomation-rs/pull/106))

## 0.19.1 - 2025-05-20

### Changed

- Updated `README.md`.

## 0.19.0 - 2025-05-19

### Changed

- Refactored the crate layout to support more features.

## 0.18.4 - 2025-05-16

### Changed

- Updated to `windows v0.61.1`.

## 0.18.2 - 2025-05-13

### Changed

- Removed the `image-data` feature from the `arboard` dependency.
  ([#105](https://github.com/leexgone/uiautomation-rs/issues/105))

## 0.18.1 - 2025-05-13

### Added

- Added `UIElement::send_text_by_clipboard()` to send text through the clipboard.
  ([#103](https://github.com/leexgone/uiautomation-rs/issues/103))

## 0.18.0 - 2025-04-10

### Added

- Added `UIMatcher::process_id()` to filter matches by process ID.

## 0.17.3 - 2025-03-29

### Added

- Made the input parser usable outside of the crate.
  ([#97](https://github.com/leexgone/uiautomation-rs/pull/97))

## 0.17.2 - 2025-03-19

### Changed

- Updated to `windows v0.61.0`.

## 0.17.1 - 2025-03-06

### Changed

- Updated dependencies.

## 0.17.0 - 2025-02-23

### Changed

- Migrated to the Rust 2024 edition.
- Updated to `windows v0.60.0`.

## 0.16.1 - 2025-02-18

### Added

- Added the `Control` derive macro.

## 0.16.0 - 2025-01-17

### Changed

- Implemented `TryFrom` for the enum types instead of `From`.

## 0.15.1 - 2025-01-15

### Changed

- Updated dependencies and documentation.

## 0.15.0 - 2025-01-12

### Added

- Added the `Keyboard::ignore_parse_err()` option to ignore parsing errors.
- Added `UIElement::send_text()` to send the literal content of a text.

## 0.14.0 - 2025-01-08

### Changed

- Updated to `windows v0.59.0`.

## 0.13.4 - 2024-12-10

### Added

- Added `UIElement::get_heading_level()`, `UIElement::get_cached_heading_level()`,
  `UIElement::is_dialog()` and `UIElement::is_cached_dialog()`.
  ([#92](https://github.com/leexgone/uiautomation-rs/pull/92))

## 0.13.3 - 2024-12-09

### Added

- Added a wrapper for `LegacyIAccessiblePattern`.
  ([#91](https://github.com/leexgone/uiautomation-rs/pull/91))

## 0.13.2 - 2024-12-02

### Changed

- Simulated mouse movement with `SendInput` and `MOUSEINPUT`.
  ([#90](https://github.com/leexgone/uiautomation-rs/pull/90))

## 0.13.1 - 2024-11-28

### Added

- Added support for the CONTROL and ALT keys.
  ([#86](https://github.com/leexgone/uiautomation-rs/pull/86))

## 0.13.0 - 2024-10-31

### Added

- Added support for the `i686-pc-windows-msvc` target.
  ([#84](https://github.com/leexgone/uiautomation-rs/issues/84))

## 0.12.3 - 2024-08-31

### Changed

- Replaced the `phf` crate with `OnceLock`.

## 0.12.2 - 2024-08-17

### Added

- Added cached variants of the `get_*`, `is_*` and `can_*` methods for all patterns.
  ([#79](https://github.com/leexgone/uiautomation-rs/pull/79))

## 0.12.1 - 2024-08-01

### Added

- Added `Window` pattern support for `PaneControl`.
  ([#75](https://github.com/leexgone/uiautomation-rs/issues/75))

## 0.12.0 - 2024-07-11

### Changed

- Updated to `windows v0.58.0`.

## 0.11.5 - 2024-07-10

### Added

- Added the `log` feature, which logs failed filters for debugging.
  ([#70](https://github.com/leexgone/uiautomation-rs/pull/70))

## 0.11.4 - 2024-06-29

### Added

- Implemented `Display` for the enum types.

## 0.11.3 - 2024-06-25

### Fixed

- Fixed `click` errors between multiple monitors.
  ([#62](https://github.com/leexgone/uiautomation-rs/issues/62))

## 0.11.2 - 2024-06-10

### Changed

- Updated to `windows v0.57.0`.

### Fixed

- Fixed the lifetime of the encoded string buffers used by `PCWSTR` and `PWSTR` pointers.
  ([#65](https://github.com/leexgone/uiautomation-rs/issues/65))

## 0.11.1 - 2024-06-07

### Added

- Added `UIElement::get_cached_children()` and `UIElement::get_cached_parent()`.
  ([#63](https://github.com/leexgone/uiautomation-rs/issues/63))

## 0.11.0 - 2024-06-01

### Added

- Added event handler support.
  ([#59](https://github.com/leexgone/uiautomation-rs/issues/59))

## 0.10.0 - 2024-04-27

### Fixed

- Fixed typos in method names (`ragne` → `range`).
  ([#57](https://github.com/leexgone/uiautomation-rs/pull/57))

## 0.9.3 - 2024-04-26

### Fixed

- Fixed the bug reported in [#52](https://github.com/leexgone/uiautomation-rs/issues/52).

## 0.9.2 - 2024-04-25

### Added

- Added `UIMatcher::from_ref()`.

## 0.9.0 - 2024-04-14

### Added

- Added cache request support.

## 0.8.3 - 2024-04-14

### Changed

- Updated to `windows v0.56.0`.

## 0.8.0 - 2024-03-14

### Changed

- Updated to `windows v0.54.0`.

## 0.7.3 - 2023-11-16

### Changed

- Updated to `windows v0.52.0`.

## 0.7.2 - 2023-11-11

### Added

- Implemented `From<isize>` for `Handle`.
  ([#44](https://github.com/leexgone/uiautomation-rs/issues/44))

## 0.7.0 - 2023-10-20

### Added

- Added more control wrappers.

### Changed

- Refactored `Process`.

## 0.6.4 - 2023-10-02

### Changed

- `Keyboard::begin_hold_keys()` and `Keyboard::end_hold_keys()` now honor the
  `Keyboard::interval()` setting.

## 0.6.0 - 2023-08-16

### Changed

- Updated to `windows v0.51.0`.

## 0.5.4 - 2023-05-16

### Changed

- Updated dependencies.

## 0.5.3 - 2023-05-11

### Changed

- Updated dependencies.

## 0.5.2 - 2023-04-24

### Changed

- Updated the sample code in `README.md`.

## 0.5.1 - 2023-04-21

### Added

- Added `UIAutomation::new_direct()` to create a `UIAutomation` without initializing COM.
  ([#37](https://github.com/leexgone/uiautomation-rs/issues/37))

## 0.5.0 - 2023-04-20

### Fixed

- Fixed a `VARIANT` array error.
  ([#35](https://github.com/leexgone/uiautomation-rs/issues/35))

## 0.4.2 - 2023-04-15 (unpublished)

### Changed

- Updated dependencies.
- This version was never published on crates.io; `0.5.0` superseded it.

## 0.4.1 - 2023-04-07

### Changed

- Updated dependencies.
- Refactored the `win-update` sample.

## 0.4.0 - 2023-04-05

### Added

- Added dialog support.

## 0.3.10 - 2023-04-01

### Changed

- Updated to `syn v2.0.12` and `windows v0.48.0`.

## 0.3.9 - 2023-03-29

### Changed

- Updated to `syn v2.0.11` and `windows v0.47.0`.

## 0.3.8 - 2023-03-27

### Changed

- Updated to `syn v2.0.10` and `proc-macro2 v1.0.54`.

## 0.3.7 - 2023-03-19

### Changed

- Updated to `syn v2.0.2`.

## 0.3.6 - 2023-03-16

### Changed

- Updated to `windows v0.46`.

## 0.3.4 - 2023-02-12

### Changed

- Updated dependencies.

## 0.3.3 - 2023-02-05

### Changed

- Refactored the enum types.

## 0.3.2 - 2023-02-04

### Changed

- Updated to `windows v0.44`.

## 0.3.1 - 2023-01-02

### Changed

- Updated dependencies.

## 0.3.0 - 2022-11-11

### Changed

- Redefined some automation types.

## 0.2.1 - 2022-10-28

### Changed

- Updated to `windows v0.43.0`.

## 0.2.0 - 2022-09-24

### Changed

- Updated to `windows v0.41.0`.
- Refactored the `UIPattern` trait.

## 0.1.12 - 2022-09-23

### Added

- Added `UIElement::hold_send_keys()` to hold keys while sending text.

## 0.1.11 - 2022-08-23

### Added

- Added `UIMatcher::filter_fn()`.

## 0.1.10 - 2022-08-03

### Added

- Added `WindowControl::set_foregrand()`.

### Changed

- Updated to `windows v0.39.0`.
