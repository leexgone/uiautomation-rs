# uiautomation-rs changelogs

## v0.1.10

+ update `windows` to `v0.39.0`
+ add `WindowCtroll.set_foregrand()`

## v0.1.11

+ add `UIMatcher.filter_fn()`

## v0.1.12

+ add `UIElement.hold_send_keys()` to support holding and send keys.

## v0.2.0

+ update to `windows-rs v0.41.0`
+ refator `UIPattern` trait

## v0.2.1

+ update to `windows-rs v0.43.0`

## v0.3.0

+ redefine some automation types

## v0.3.1

+ update dependencies

## v0.3.2

+ update to windwos v0.44

## v0.3.3

+ refactor enum types

## v0.3.4

+ update dependencies

## v0.3.6

+ update to `windows v0.46`

## v0.3.7

+ update to `syn v2.0.2`

## v0.3.8

+ update to `syn v2.0.10` & `proc-macro2 v1.0.54`

## v0.3.9

+ update to `syn v2.0.11` & `windows v0.47.0`

## v0.3.10

+ update to `syn v2.0.12` & `windows v0.48.0`

## v0.4.0

+ add dialogs support.

## v0.4.1

+ update dependencies
+ refactor `win-update` sample codes.

## v0.4.2

+ update dependencies

## v0.5.0

+ fix variant array error. [issue-35](https://github.com/leexgone/uiautomation-rs/issues/35)

## v0.5.1

+ add `UIAutomation::new_direct()` to create `UIAutomation` without initializing the COM library. [issue-37](https://github.com/leexgone/uiautomation-rs/issues/37)

## v0.5.2

+ Update sample codes in README

## v0.5.3

+ update dependencies

## v0.5.4

+ update dependencies

## v0.6.0

+ update to `windows v0.51.0`

## v0.6.4

+ `Keyboard.begin_hold_keys()` & `Keyboard.end_hold_keys()` respond to `Keyboard.interval()` setting.

## v0.7.0

+ refactor `Process`, add more controls.

## v0.7.2

+ support converting `isize` to `Handle`. [issue-44](https://github.com/leexgone/uiautomation-rs/issues/44)

## v0.7.3

+ update to `widnows v0.52.0`

## v0.8.0

+ update to `windows v0.54.0`

## v0.8.3

+ update to `windows v0.56.0`

## v0.9.0

+ support cache request.

## v0.9.2

+ addd `UIMatcher::from_ref()`

## v0.9.3

+ fix [issues-52](https://github.com/leexgone/uiautomation-rs/issues/52)

## v0.10.0

+ Fix typo in method names (ragne -> range) [#57](https://github.com/leexgone/uiautomation-rs/pull/57)

## v0.11.0

+ Support event handler. [#59](https://github.com/leexgone/uiautomation-rs/issues/59)

## v0.11.1

+ Support `UIElement::get_cached_children()` & `UIElement::get_cached_parent()`. [#63](https://github.com/leexgone/uiautomation-rs/issues/63)

## v0.11.2

+ Cache encoded string buffers so that they outlive PCWSTR and PWSTR pointers. [#65](https://github.com/leexgone/uiautomation-rs/issues/65)
+ Update to `windows v0.57.0`

## v0.11.3

+ fixed `click` error between multi-monitors. [#62](https://github.com/leexgone/uiautomation-rs/issues/62)

## v0.11.4

+ enum types support `Dispaly` trait.

## v0.11.5

+ Add log feature, show failed filter for debugging [#70](https://github.com/leexgone/uiautomation-rs/pull/70)

## v0.12.0

+ update to `windows v0.58.0`

## v0.12.1

+ Support `Window` pattern for `PaneControl` [#75](https://github.com/leexgone/uiautomation-rs/issues/75)

## v0.12.2

+ Add cached versions for "get_xxx"/"is_xxx"/"can_xxx" for all patterns. [#79](https://github.com/leexgone/uiautomation-rs/pull/79)

## v0.12.3

+ Use `OnceLock` instead of `phf` crate.

## v0.13.0

+ Support `i686-pc-windows-msvc` target. [#84](https://github.com/leexgone/uiautomation-rs/issues/84)

## v0.13.1

+ Add support for pressing CONTROL and ALT keys. [#86](https://github.com/leexgone/uiautomation-rs/pull/86)

## v0.13.2

+ Use SendInput({ MOUSEINPUT { ... } }) for simulating mouse movement. [#90](https://github.com/leexgone/uiautomation-rs/pull/90)

## v0.13.3

+ Add a wrapper for LegacyIAccessiblePattern. [#91](https://github.com/leexgone/uiautomation-rs/pull/91)

## v0.13.4

+ Add get_(cached)_heading_level and is_(cached)_dialog. [#92](https://github.com/leexgone/uiautomation-rs/pull/92)

## v0.14.0

+ Update to `windows v0.59.0`.

## v0.15.0

+ Add `Keyboard.ignore_parse_err()` option to ignore parsing error.
+ Add `UIElement.send_text()` method to output the literal content of the text.

## v0.15.1

+ Update deps & append docs.

## v0.16.0

+ Impl `TryFrom` for enum types, instead of `From`.

## v0.16.1

+ Add `Control` derive.

## v0.17.0

+ Migrating to `Rust edition 2024`.
+ Update to `windows v0.60.0`.

## v0.17.1

+ Update deps.

## v0.17.2

+ update to `windows v0.61.0`.

## v0.17.3

+ Allow using the input parser outside of `uiautomation-rs`. [#97](https://github.com/leexgone/uiautomation-rs/pull/97)

## v0.18.0

+ Add `UIMatcher.process_id()` to support filter by process id.

## v0.18.1

+ Add `UIElement.send_text_by_clipboard()` to support sending text by clipboard. [#103](https://github.com/leexgone/uiautomation-rs/issues/103)

## v0.18.2

+ Remove `image-data` feature from `arboard` crate. [#105](https://github.com/leexgone/uiautomation-rs/issues/105)

## v0.18.4

+ Update to `windows v0.61.1`.

## v0.19.0

+ Refactor to support more features.

## v0.19.1

+ Update `README.md`.

## v0.19.2

+ Add more key names. [#106](https://github.com/leexgone/uiautomation-rs/pull/106)

## v0.20.0

+ Refactor `process`.

## v0.21.0

+ Add `clipboard` feature, support windows clipboard.

## v0.21.1

+ Update to `windows v0.61.2`

## v0.21.2

+ Change docs settings.

## v0.21.3

+ Update to `windows v0.61.3`
