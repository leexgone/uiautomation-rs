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
