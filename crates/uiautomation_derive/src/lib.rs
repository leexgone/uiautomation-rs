extern crate proc_macro;

#[cfg(feature = "enum_derive")]
mod enum_derives;
#[cfg(feature = "action_derive")]
mod action_derives;
#[cfg(feature = "control_derive")]
mod control_derives;
#[cfg(feature = "pattern_derive")]
mod pattern_derivers;

use proc_macro::TokenStream;
use syn::ItemEnum;
use syn::Path;

#[cfg(feature = "enum_derive")]
use self::enum_derives::*;
#[cfg(feature = "action_derive")]
use self::action_derives::*;
#[cfg(feature = "control_derive")]
use self::control_derives::*;
#[cfg(feature = "pattern_derive")]
use self::pattern_derivers::*;

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Invoke)]
pub fn derive_invoke(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_invoke(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(SelectionItem)]
pub fn derive_selection_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_selection_item(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(MultipleView)]
pub fn derive_multiple_view(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_multiple_view(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(ItemContainer)]
pub fn item_container_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_item_container(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(ScrollItem)]
pub fn derive_scroll_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_scroll_item(&ast)    
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Window)]
pub fn derive_window(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_window(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Transform)]
pub fn derive_transform(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_transform(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Value)]
pub fn derive_value(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_value(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(ExpandCollapse)]
pub fn derive_expand_collapse(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_expand_collapse(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Toggle)]
pub fn derive_toggle(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_toggle(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Grid)]
pub fn derive_grid(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_grid(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Table)]
pub fn derive_table(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();
    
    impl_table(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Scroll)]
pub fn derive_scroll(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_scroll(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Selection)]
pub fn derive_selection(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_selection(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(CustomNavigation)]
pub fn derive_custom_navigation(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_custom_navigation(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(GridItem)]
pub fn derive_grid_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_grid_item(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(TableItem)]
pub fn derive_table_item(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_table_item(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(Text)]
pub fn derive_text(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_text(&ast)
}

#[cfg(feature = "action_derive")]
#[proc_macro_derive(RangeValue)]
pub fn derive_range_value(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_range_value(&ast)    
}

#[cfg(feature = "action_derive")]
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

#[cfg(feature = "control_derive")]
#[proc_macro_derive(Control)]
pub fn derive_control(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_control(&ast)
}

#[proc_macro_attribute]
pub fn map_as(args: TokenStream, item: TokenStream) -> TokenStream {
    let type_path: Path = syn::parse(args).expect("#[maps_as() requires type path");
    let enum_item: ItemEnum = syn::parse(item).expect("#[map_as()] must be used on enums only");

    impl_map_as(type_path, enum_item)
}

#[cfg(feature = "pattern_derive")]
#[proc_macro_attribute]
pub fn pattern_as(attr: TokenStream, item: TokenStream) -> TokenStream {
    use syn::punctuated::Punctuated;
    use syn::ItemStruct;
    use syn::Token;

    let struct_item: ItemStruct = syn::parse(item).expect("#[pattern_as()] must be used on struct only");

    let parser = Punctuated::<syn::Path, Token![,]>::parse_terminated;
    let args = syn::parse_macro_input!(attr with parser);
    // let args: Vec<_> = types.into_iter().collect();
    // if args.len() != 2 {
    //     syn::Error::new(
    //         proc_macro2::Span::call_site(), 
    //         "#[pattern_as()] requires both pattern type and interface"
    //     ).to_compile_error()
    //     .into()
    // } else {
    //     todo!()
    // }
    let mut arg_itr = args.into_iter();
    let type_item = arg_itr.next().expect("#[pattern_as()] requires pattern type");
    // let pattern_item = arg_itr.next().expect("#[pattern_as()] requires pattern interface");
    if arg_itr.next().is_some() {
        return syn::Error::new(
            proc_macro2::Span::call_site(), 
            "#[pattern_as()] support only single parameter"
        ).to_compile_error()
        .into();
    }

    impl_pattern_as(struct_item, type_item)
}