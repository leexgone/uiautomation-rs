extern crate proc_macro;

mod action_derives;

use proc_macro::TokenStream;

use crate::action_derives::impl_invoke;
use crate::action_derives::impl_selection_item;

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
