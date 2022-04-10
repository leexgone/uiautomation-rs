extern crate proc_macro;

mod control_derives;

use proc_macro::TokenStream;

use crate::control_derives::impl_click;
use crate::control_derives::impl_select;

#[proc_macro_derive(Click)]
pub fn click_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_click(&ast)
}

#[proc_macro_derive(Select)]
pub fn select_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_select(&ast)
}
