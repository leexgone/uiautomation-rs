use proc_macro::TokenStream;
use proc_macro2::Ident;
use proc_macro2::TokenTree;
use quote::format_ident;
use quote::quote;
use syn::AttributeArgs;
use syn::Expr;
use syn::ItemEnum;
use syn::Meta;
use syn::NestedMeta;
use syn::parse_quote;

const REPR_TYPES: &[&'static str] = &["u8", "u16", "u32", "u64", "usize", "i8", "i16", "i32", "i64", "isize"];

pub(crate) fn impl_enum_convert(enum_item: ItemEnum) -> TokenStream {
    let enum_name = &enum_item.ident;
    let enum_type = get_repr_type(&enum_item).expect(&format!("#[EnumConvert] support #[repr({})] only", REPR_TYPES.join(" | ")));
    let (enum_names, var_exprs) = get_variants(&enum_item);

    let enum_name_upper = enum_name.to_string().to_uppercase();
    let var_names: Vec<Ident> = enum_names.iter().map(|n| {
        format_ident!("_{}_{}_", enum_name_upper, n.to_string().to_uppercase())
    }).collect();

    let gen = quote! {
        impl TryFrom<#enum_type> for #enum_name {
            type Error = crate::errors::Error;
        
            fn try_from(value: #enum_type) -> core::result::Result<Self, Self::Error> {
                #( const #var_names: #enum_type = #var_exprs; )*

                match value {
                    #( #var_names => Ok(Self::#enum_names), )*
                    _ => Err(crate::errors::Error::new(crate::errors::ERR_NOTFOUND, "Unsupported value"))
                }
            }
        }        
    };

    gen.into()
}

fn get_repr_type(enum_item: &syn::ItemEnum) -> Option<Ident> {
    let repr_attr = enum_item.attrs.iter().filter(|a| {
        (*a).path.is_ident("repr")
        // if let Some(ident) = (*a).path.get_ident() {
        //     ident == "repr"
        // } else {
        //     false
        // }
    }).next().expect("#[EnumConvert] must be used on enum which has #[repr(_)]");

    let tokens = repr_attr.tokens.clone();
    if let Some(TokenTree::Group(group)) = tokens.into_iter().next() {
        group.stream().into_iter().filter_map(|t| {
            if let TokenTree::Ident(ident) = t {
                let type_name = ident.to_string();
                if REPR_TYPES.iter().any(|t| *t == &type_name) {
                    Some(ident)
                } else {
                    None
                }
            } else {
                None
            }
        }).next()
    } else {
        None
    }
}

fn get_variants(enum_item: &ItemEnum) -> (Vec<Ident>, Vec<Expr>) {
    let mut prev_expr: Option<Expr> = None;
    enum_item.variants.iter().map(|v| {
        let expr = if let Some((_, ref expr)) = v.discriminant {
            expr.clone()
        } else {
            if let Some(ref prev_expr) = prev_expr {
                parse_quote!(#prev_expr + 1)
            } else {
                parse_quote!(0)
            }
        };
        prev_expr = Some(expr.clone());

        (v.ident.clone(), expr)
    }).unzip()
}

pub(crate) fn impl_map_as(args: AttributeArgs, enum_item: ItemEnum) -> TokenStream {
    let enum_name = &enum_item.ident;

    let mut gen = quote! {
        #enum_item

    };

    for arg in args.into_iter() {
        if let NestedMeta::Meta(Meta::Path(type_path)) = arg {
            let mapping = quote!{
                impl From<#type_path> for #enum_name {
                    fn from(value: #type_path) -> Self {
                        value.0.try_into().unwrap()
                    }
                }
                
                impl Into<#type_path> for #enum_name {
                    fn into(self) -> #type_path {
                        #type_path(self as _)
                    }
                }        
            };

            gen.extend(mapping);
        }
    }

    gen.into()
}