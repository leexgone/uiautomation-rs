use proc_macro::TokenStream;
use quote::quote;

pub(crate) fn impl_control(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;

    let gen = quote! {
        impl Into<UIElement> for #name {
            fn into(self) -> UIElement {
                self.control
            }
        }
    };

    gen.into()
}