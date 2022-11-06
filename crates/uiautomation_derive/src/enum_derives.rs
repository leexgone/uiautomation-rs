use proc_macro::TokenStream;
use proc_macro2::Ident;
use proc_macro2::TokenTree;
use quote::quote;
use syn::Expr;
use syn::ItemEnum;
use syn::parse_quote;

const REPR_TYPES: &[&'static str] = &["u8", "u16", "u32", "u64", "usize", "i8", "i16", "i32", "i64", "isize"];

pub(crate) fn impl_enum_convert(enum_item: &ItemEnum) -> TokenStream {
    let enum_name = &enum_item.ident;
    let enum_type = get_repr_type(enum_item).expect(&format!("#[EnumConvert] support #[repr({})] only", REPR_TYPES.join(" | ")));
    let (var_names, var_exprs) = get_variants(enum_item);
    
    let gen = quote! {
        impl TryFrom<#enum_type> for #enum_name {
            type Error = crate::errors::Error;
        
            fn try_from(value: #enum_type) -> Result<Self, Self::Error> {
                match value {
                    // 0i32 => Ok(WindowInteractionState::Running),
                    // 1i32 => Ok(WindowInteractionState::Closing),
                    // 2i32 => Ok(WindowInteractionState::ReadyForUserInteraction),
                    // 3i32 => Ok(WindowInteractionState::BlockedByModalWindow),
                    // 4i32 => Ok(WindowInteractionState::NotResponding),
                    #( #var_exprs => Ok(Self::#var_names), )*
                    _ => Err(crate::errors::Error::new(crate::errors::ERR_NOTFOUND, "Unsupported value"))
                }
            }
        }        
    };

    gen.into()
}

fn get_repr_type(enum_item: &syn::ItemEnum) -> Option<Ident> {
    let repr_attr = enum_item.attrs.iter().filter(|a| {
        if let Some(ident) = (*a).path.get_ident() {
            ident == "repr"
        } else {
            false
        }
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
