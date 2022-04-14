use proc_macro::TokenStream;
use quote::quote;

pub(crate) fn impl_invoke(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Invoke for #name {
            fn invoke(&self) -> Result<()> {
                let pattern: UIInvokePattern = self.as_ref().get_pattern()?;
                pattern.invoke()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_selection_item(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl SelectionItem for #name {
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

pub(crate) fn impl_multiple_view(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl MultipleView for #name {
            fn get_supported_views(&self) -> Result<Vec<i32>> {
                let pattern: UIMultipleViewPattern = self.as_ref().get_pattern()?;
                pattern.get_supported_views()
            }

            fn get_view_name(&self, view: i32) -> Result<String> {
                let pattern: UIMultipleViewPattern = self.as_ref().get_pattern()?;
                pattern.get_view_name(view)
            }

            fn get_current_view(&self) -> Result<i32> {
                let pattern: UIMultipleViewPattern = self.as_ref().get_pattern()?;
                pattern.get_current_view()
            }

            fn set_current_view(&self, view: i32) -> Result<()> {
                let pattern: UIMultipleViewPattern = self.as_ref().get_pattern()?;
                pattern.set_current_view(view)
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_item_container(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl ItemContainer for #name {
            fn find_item_by_property(&self, start_after: UIElement, property_id: i32, value: Variant) -> Result<UIElement> {
                let pattern: UIItemContainerPattern = self.as_ref().get_pattern()?;
                pattern.find_item_by_property(start_after, property_id, value)
            }        
        }
    };
    gen.into()
}

pub(crate) fn impl_scroll_item(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl ScrollItem for #name {
            fn scroll_into_view(&self) -> Result<()> {
                let pattern: UIScrollItemPattern = self.as_ref().get_pattern()?;
                pattern.scroll_into_view()
            }
        }
    };
    gen.into()
}