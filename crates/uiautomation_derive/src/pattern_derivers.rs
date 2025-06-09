use proc_macro::TokenStream;
use quote::quote;
use syn::Fields;
use syn::ItemStruct;
use syn::Path;
use syn::Type;

pub(crate) fn impl_pattern_as(struct_item: ItemStruct, type_item: Path) -> TokenStream {
    let struct_name = &struct_item.ident;
    let pattern_type = get_pattern_type(&struct_item).expect("#[pattern_as()] can't find pattern type in struct");

    let (inherit_names, inherit_types) = get_pattern_inherits(&struct_item);
    
    quote! {
        #struct_item

        impl UIPattern for #struct_name {
            const TYPE: UIPatternType = #type_item;
        }

        impl From<#pattern_type> for #struct_name {
            fn from(pattern: #pattern_type) -> Self {
                #(let #inherit_names: windows::core::IUnknown = pattern.cast().unwrap();)*
                Self {
                    pattern,
                    #(#inherit_names: #inherit_names.try_into().unwrap(),)*
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

        #(
            impl AsRef<#inherit_types> for #struct_name {
                fn as_ref(&self) -> &#inherit_types {
                    &self.#inherit_names
                }
            }

            impl Into<#inherit_types> for #struct_name {
                fn into(self) -> #inherit_types {
                    self.#inherit_names
                }
            }
        )*   
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

fn get_pattern_inherits<'a>(struct_item: &'a ItemStruct) -> (Vec<&'a proc_macro2::Ident>, Vec<&'a Type>) {
    if let Fields::Named(fields) = &struct_item.fields {
        fields.named.iter().filter_map(|field| {
            if let Some(ref field_name) = field.ident {
                if field_name != "pattern" {
                    Some((field_name, &field.ty))
                } else {
                    None
                }
            } else {
                None
            }
        }).unzip()
    } else {
        (Vec::new(), Vec::new())
    }
}
