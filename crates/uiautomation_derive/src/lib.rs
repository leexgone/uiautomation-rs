use proc_macro::TokenStream;
use quote::quote;

extern crate proc_macro;

#[proc_macro_derive(Click)]
pub fn click_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_click(&ast)
}

fn impl_click(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Click for #name {
            fn click(&self) -> Result<()> {
                let pattern: UIInvokePattern = self.as_ref().get_pattern()?;
                pattern.invoke()
            }
        }
    };
    gen.into()
}

#[proc_macro_derive(Select)]
pub fn select_derive(input: TokenStream) -> TokenStream {
    let ast = syn::parse(input).unwrap();

    impl_select(&ast)
}

fn impl_select(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Select for #name {
            fn select(&self) -> Result<()> {
                let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
                pattern.select()
            }

            fn add_to_selection(&self) -> Result<()> {
                let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
                pattern.add_to_selection()
            }

            fn remove_from_selection(&self) -> Result<()> {
                let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
                pattern.remove_from_selection()
            }

            fn is_selected(&self) -> Result<bool> {
                let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
                pattern.is_selected()
            }
        }
    };
    gen.into()
}