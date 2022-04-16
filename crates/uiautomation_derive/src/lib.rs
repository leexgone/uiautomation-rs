extern crate proc_macro;

mod action_derives;

use proc_macro::TokenStream;

use crate::action_derives::impl_invoke;
use crate::action_derives::impl_item_container;
use crate::action_derives::impl_multiple_view;
use crate::action_derives::impl_scroll_item;
use crate::action_derives::impl_selection_item;
use crate::action_derives::impl_transform;
use crate::action_derives::impl_window;

#[proc_macro_derive(Invoke)]
pub fn invoke_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_invoke(&ast)
}

#[proc_macro_derive(SelectionItem)]
pub fn selection_item_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_selection_item(&ast)
}

#[proc_macro_derive(MultipleView)]
pub fn multiple_view_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_multiple_view(&ast)
}

#[proc_macro_derive(ItemContainer)]
pub fn item_container_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_item_container(&ast)
}

#[proc_macro_derive(ScrollItem)]
pub fn scroll_item_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_scroll_item(&ast)    
}

#[proc_macro_derive(Window)]
pub fn window_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_window(&ast)
}

#[proc_macro_derive(Transform)]
pub fn derive_transform(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_transform(&ast)
}