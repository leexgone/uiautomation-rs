use proc_macro::TokenStream;
use quote::quote;
use syn::Fields;
use syn::ItemStruct;
use syn::Path;
use syn::Type;

pub(crate) fn impl_pattern_as(struct_item: ItemStruct, type_item: Path) -> TokenStream {
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
    let pattern_type = get_pattern_type(&struct_item).expect("#[pattern_as()] can't find pattern type in struct");
    
    quote! {
        #struct_item

        impl UIPattern for #struct_name {
            const TYPE: UIPatternType = #type_item;
        }

        impl From<#pattern_type> for #struct_name {
            fn from(pattern: #pattern_type) -> Self {
                Self {
                    pattern
                }
            }
        }

        impl Into<#pattern_type> for #struct_name {
            fn into(self) -> #pattern_type {
                self.pattern
            }
        }

        impl AsRef<#pattern_type> for #struct_name {
            fn as_ref(&self) -> &#pattern_type {
                &self.pattern
            }
        }

        impl TryFrom<IUnknown> for #struct_name {
            type Error = super::errors::Error;

            fn try_from(pattern: IUnknown) -> core::result::Result<Self, Self::Error> {
                let pattern: #pattern_type = pattern.cast()?;
                Ok(pattern.into())
            }
        }        
    }.into()
}

fn get_pattern_type<'a>(struct_item: &'a ItemStruct) -> Option<&'a Type> {
    if let Fields::Named(fields) = &struct_item.fields {
        fields.named.iter().find_map(|field| {
            if let Some(ref field_name) = field.ident {
                if field_name == "pattern" {
                    Some(&field.ty)
                } else {
                    None
                }
            } else {
                None
            }
        })
    } else {
        None
    }
}
