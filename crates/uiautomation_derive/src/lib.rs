extern crate proc_macro;

mod action_derives;

use proc_macro::TokenStream;

use crate::action_derives::*;

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