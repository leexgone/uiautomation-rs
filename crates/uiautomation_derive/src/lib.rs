extern crate proc_macro;

mod action_derives;

use proc_macro::TokenStream;

use crate::action_derives::impl_invoke;
use crate::action_derives::impl_item_container;
use crate::action_derives::impl_multiple_view;
use crate::action_derives::impl_scroll_item;
use crate::action_derives::impl_selection_item;
use crate::action_derives::impl_transform;
use crate::action_derives::impl_value;
use crate::action_derives::impl_window;

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