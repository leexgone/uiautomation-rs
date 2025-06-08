use proc_macro::TokenStream;
use quote::quote;
use syn::ItemStruct;
use syn::Path;

pub(crate) fn impl_pattern_as(struct_item: ItemStruct, type_item: Path, pattern_item: Path) -> TokenStream {
    // let enum_name = &enum_item.ident;

    // quote! {
    //     #enum_item

    //     impl TryFrom<#type_path> for #enum_name {
    //         type Error = super::errors::Error;
    //         fn try_from(value: #type_path) -> super::errors::Result<Self> {
    //             value.0.try_into()
    //         }
    //     }

    //     impl Into<#type_path> for #enum_name {
    //         fn into(self) -> #type_path {
    //             #type_path(self as _)
    //         }
    //     }        
    // }.into()
    let struct_name = &struct_item.ident;
    
    quote! {
        #struct_item

        impl UIPattern for #struct_name {
            const TYPE: UIPatternType = #type_item;
        }

        impl From<#pattern_item> for #struct_name {
            fn from(pattern: #pattern_item) -> Self {
                Self {
                    pattern
                }
            }
        }

        impl Into<#pattern_item> for #struct_name {
            fn into(self) -> #pattern_item {
                self.pattern
            }
        }

        impl AsRef<#pattern_item> for #struct_name {
            fn as_ref(&self) -> &#pattern_item {
                &self.pattern
            }
        }

        impl TryFrom<IUnknown> for #struct_name {
            type Error = super::errors::Error;

            fn try_from(pattern: IUnknown) -> core::result::Result<Self, Self::Error> {
                let pattern: #pattern_item = pattern.cast()?;
                // Ok(Self {
                //     pattern
                // })
                Ok(pattern.into())
            }
        }        
    }.into()
}
