extern crate proc_macro;

mod action_derives;
mod enum_derives;

use proc_macro::TokenStream;
use syn::ItemEnum;
use syn::Path;

use self::action_derives::*;
use self::enum_derives::*;

#[proc_macro_derive(Invoke)]
pub fn derive_invoke(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_invoke(&ast)
}

#[proc_macro_derive(SelectionItem)]
pub fn derive_selection_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_selection_item(&ast)
}

#[proc_macro_derive(MultipleView)]
pub fn derive_multiple_view(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_multiple_view(&ast)
}

#[proc_macro_derive(ItemContainer)]
pub fn item_container_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_item_container(&ast)
}

#[proc_macro_derive(ScrollItem)]
pub fn derive_scroll_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_scroll_item(&ast)    
}

#[proc_macro_derive(Window)]
pub fn derive_window(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_window(&ast)
}

#[proc_macro_derive(Transform)]
pub fn derive_transform(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_transform(&ast)
}

#[proc_macro_derive(Value)]
pub fn derive_value(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_value(&ast)
}

#[proc_macro_derive(ExpandCollapse)]
pub fn derive_expand_collapse(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_expand_collapse(&ast)
}

#[proc_macro_derive(Toggle)]
pub fn derive_toggle(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_toggle(&ast)
}

#[proc_macro_derive(Grid)]
pub fn derive_grid(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_grid(&ast)
}

#[proc_macro_derive(Table)]
pub fn derive_table(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();
    
    impl_table(&ast)
}

#[proc_macro_derive(Scroll)]
pub fn derive_scroll(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_scroll(&ast)
}

#[proc_macro_derive(Selection)]
pub fn derive_selection(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_selection(&ast)
}

#[proc_macro_derive(CustomNavigation)]
pub fn derive_custom_navigation(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_custom_navigation(&ast)
}

#[proc_macro_derive(GridItem)]
pub fn derive_grid_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_grid_item(&ast)
}

#[proc_macro_derive(TableItem)]
pub fn derive_table_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_table_item(&ast)
}

#[proc_macro_derive(Text)]
pub fn derive_text(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_text(&ast)
}

#[proc_macro_derive(RangeValue)]
pub fn derive_range_value(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_range_value(&ast)    
}

#[proc_macro_derive(Dock)]
pub fn derive_dock(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_dock(&ast)    
}

#[proc_macro_derive(EnumConvert)]
pub fn derive_enum_convert(input: TokenStream) -> TokenStream {
    let enum_item = syn::parse::<ItemEnum>(input).expect("#[Derive(EnumConvert)] must be used on enums only");

    impl_enum_convert(enum_item)
}

#[proc_macro_attribute]
pub fn map_as(args: TokenStream, item: TokenStream) -> TokenStream {
    let enum_item: ItemEnum = syn::parse(item).expect("#[map_as()] must be used on enums only");
    let type_path: Path = syn::parse(args).expect("#[maps_as() requires type path");

    impl_map_as(type_path, enum_item)
}
