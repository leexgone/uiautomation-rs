use proc_macro::TokenStream;
use quote::quote;

pub(crate) fn impl_control(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;

    let gen = quote! {
        impl TryFrom<UIElement> for #name {
            type Error = super::errors::Error;

            fn try_from(value: UIElement) -> super::errors::Result<Self> {
                if value.get_control_type()? == Self::TYPE {
                    Ok(Self {
                        control: value
                    })
                } else {
                    Err(super::errors::Error::new(ERR_TYPE, "Error Control Type"))
                }
            }
        }

        impl TryFrom<&UIElement> for #name {
            type Error = super::errors::Error;

            fn try_from(value: &UIElement) -> super::errors::Result<Self> {
                if value.get_control_type()? == Self::TYPE {
                    Ok(Self {
                        control: value.clone()
                    })
                } else {
                    Err(super::errors::Error::new(ERR_TYPE, "Error Control Type"))
                }
            }
        }

        impl Into<UIElement> for #name {
            fn into(self) -> UIElement {
                self.control
            }
        }

        impl AsRef<UIElement> for #name {
            fn as_ref(&self) -> &UIElement {
                &self.control
            }
        }

        impl Display for #name {
            fn fmt(&self, f: &mut Formatter) -> fmt::Result {
                write!(f, "{}({})", Self::TYPE, self.control.get_name().unwrap_or_default())
            }
        }
    };

    gen.into()
}