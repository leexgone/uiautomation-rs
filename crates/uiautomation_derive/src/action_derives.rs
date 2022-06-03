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

            fn get_selection_container(&self) -> Result<UIElement> {
                let pattern: UISelectionItemPattern = self.as_ref().get_pattern()?;
                pattern.get_selection_container()
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

pub(crate) fn impl_window(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Window for #name {
            fn close(&self) -> Result<()> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.close()
            }
        
            fn wait_for_input_idle(&self, milliseconds: i32) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.wait_for_input_idle(milliseconds)
            }
        
            fn is_normal(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                Ok(pattern.get_window_visual_state()? == windows::Win32::UI::Accessibility::WindowVisualState_Normal)
            }
        
            fn normal(&self) -> Result<()> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.set_window_visual_state(windows::Win32::UI::Accessibility::WindowVisualState_Normal)
            }
        
            fn can_maximize(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.can_maximize()
            }
        
            fn is_maximized(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                Ok(pattern.get_window_visual_state()? == windows::Win32::UI::Accessibility::WindowVisualState_Maximized)
            }
        
            fn maximize(&self) -> Result<()> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.set_window_visual_state(windows::Win32::UI::Accessibility::WindowVisualState_Maximized)
            }
        
            fn can_minimize(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.can_minimize()
            }
        
            fn is_minimized(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                Ok(pattern.get_window_visual_state()? == windows::Win32::UI::Accessibility::WindowVisualState_Minimized)
            }
        
            fn minimize(&self) -> Result<()> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.set_window_visual_state(windows::Win32::UI::Accessibility::WindowVisualState_Minimized)
            }
        
            fn is_modal(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.is_modal()
            }
        
            fn is_topmost(&self) -> Result<bool> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.is_topmost()
            }
        
            fn get_window_interaction_state(&self) -> Result<windows::Win32::UI::Accessibility::WindowInteractionState> {
                let pattern: UIWindowPattern = self.as_ref().get_pattern()?;
                pattern.get_window_interaction_state()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_transform(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Transform for #name {
            fn can_move(&self) -> Result<bool> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.can_move()
            }
        
            fn move_to(&self, x: f64, y: f64) -> Result<()> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.move_to(x, y)
            }
        
            fn can_resize(&self) -> Result<bool> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.can_resize()
            }
        
            fn resize(&self, width: f64, height: f64) -> Result<()> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.resize(width, height)
            }
        
            fn can_rotate(&self) -> Result<bool> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.can_rotate()
            }
        
            fn rotate(&self, degrees: f64) -> Result<()> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.rotate(degrees)
            }
        
            fn can_zoom(&self) -> Result<bool> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.can_zoom()
            }
        
            fn get_zoom_level(&self) -> Result<f64> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.get_zoom_level()
            }
        
            fn get_zoom_minimum(&self) -> Result<f64> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.get_zoom_minimum()
            }
        
            fn get_zoom_maximum(&self) -> Result<f64> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.get_zoom_maximum()
            }
        
            fn zoom(&self, zoom_value: f64) -> Result<()> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.zoom(zoom_value)
            }
        
            fn zoom_by_unit(&self, zoom_unit: windows::Win32::UI::Accessibility::ZoomUnit) -> Result<()> {
                let pattern: UITransformPattern = self.as_ref().get_pattern()?;
                pattern.zoom_by_unit(zoom_unit)
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_value(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Value for #name {
            fn set_value(&self, value: &str) -> Result<()> {
                let pattern: UIValuePattern = self.as_ref().get_pattern()?;
                pattern.set_value(value)
            }
        
            fn get_value(&self) -> Result<String> {
                let pattern: UIValuePattern = self.as_ref().get_pattern()?;
                pattern.get_value()
            }
        
            fn is_readonly(&self) -> Result<bool> {
                let pattern: UIValuePattern = self.as_ref().get_pattern()?;
                pattern.is_readonly()
            }
        }        
    };
    gen.into()
}

pub(crate) fn impl_expand_collapse(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl ExpandCollapse for #name {
            fn expand(&self) -> Result<()> {
                let pattern: UIExpandCollapsePattern = self.as_ref().get_pattern()?;
                pattern.expand()
            }
        
            fn collapse(&self) -> Result<()> {
                let pattern: UIExpandCollapsePattern = self.as_ref().get_pattern()?;
                pattern.collapse()
            }
        
            fn get_state(&self) -> Result<ExpandCollapseState> {
                let pattern: UIExpandCollapsePattern = self.as_ref().get_pattern()?;
                pattern.get_state()
            }
        }
            };
    gen.into()
}

pub(crate) fn impl_toggle(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Toggle for #name {
            fn get_toggle_state(&self) -> Result<ToggleState> {
                let pattern: UITogglePattern = self.as_ref().get_pattern()?;
                pattern.get_toggle_state()
            }
        
            fn toggle(&self) -> Result<()> {
                let pattern: UITogglePattern = self.as_ref().get_pattern()?;
                pattern.toggle()
            }
        }
    };
    gen.into()    
}

pub(crate) fn impl_grid(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Grid for #name {
            fn get_column_count(&self) -> Result<i32> {
                let pattern: UIGridPattern = self.as_ref().get_pattern()?;
                pattern.get_column_count()
            }
        
            fn get_row_count(&self) -> Result<i32> {
                let pattern: UIGridPattern = self.as_ref().get_pattern()?;
                pattern.get_row_count()
            }
        
            fn get_item(&self, row: i32, column: i32) -> Result<UIElement> {
                let pattern: UIGridPattern = self.as_ref().get_pattern()?;
                pattern.get_item(row, column)
            }
        }        
    };
    gen.into()    
}

pub(crate) fn impl_table(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Table for #name {
            fn get_row_headers(&self) -> Result<Vec<UIElement>> {
                let pattern: UITablePattern = self.as_ref().get_pattern()?;
                pattern.get_row_headers()
            }
        
            fn get_column_headers(&self) -> Result<Vec<UIElement>> {
                let pattern: UITablePattern = self.as_ref().get_pattern()?;
                pattern.get_column_headers()
            }
        
            fn get_row_or_column_major(&self) -> Result<RowOrColumnMajor> {
                let pattern: UITablePattern = self.as_ref().get_pattern()?;
                pattern.get_row_or_column_major()
            }
        }        
    };
    gen.into()    
}

pub(crate) fn impl_scroll(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Scroll for #name {
            fn scroll(&self, horizontal_amount: ScrollAmount, vertical_amount: ScrollAmount) -> Result<()> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.scroll(horizontal_amount, vertical_amount)
            }
        
            fn set_scroll_percent(&self, horizontal_percent: f64, vertical_percent: f64) -> Result<()> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.set_scroll_percent(horizontal_percent, vertical_percent)
            }
        
            fn get_horizontal_scroll_percent(&self) -> Result<f64> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.get_horizontal_scroll_percent()
            }
        
            fn get_vertical_scroll_percent(&self) -> Result<f64> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.get_vertical_scroll_percent()
            }
        
            fn get_horizontal_view_size(&self) -> Result<f64> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.get_horizontal_view_size()
            }
        
            fn get_vertical_view_size(&self) -> Result<f64> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.get_vertical_view_size()
            }
        
            fn is_horizontally_scrollable(&self) -> Result<bool> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.is_horizontally_scrollable()
            }
        
            fn is_vertically_scrollable(&self) -> Result<bool> {
                let pattern: UIScrollPattern = self.as_ref().get_pattern()?;
                pattern.is_vertically_scrollable()
            }
        }
    };
    gen.into()    
}

pub(crate) fn impl_selection(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Selection for #name {
            fn get_selection(&self) -> Result<Vec<UIElement>> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.get_selection()
            }
        
            fn can_select_multiple(&self) -> Result<bool> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.can_select_multiple()
            }
        
            fn is_selection_required(&self) -> Result<bool> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.is_selection_required()
            }
        
            fn get_first_selected_item(&self) -> Result<UIElement> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.get_first_selected_item()
            }
        
            fn get_last_selected_item(&self) -> Result<UIElement> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.get_last_selected_item()
            }
        
            fn get_current_selected_item(&self) -> Result<UIElement> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.get_current_selected_item()
            }
        
            fn get_item_count(&self) -> Result<i32> {
                let pattern: UISelectionPattern = self.as_ref().get_pattern()?;
                pattern.get_item_count()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_custom_navigation(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl CustomNavigation for #name {
            fn navigate(&self, direction: NavigateDirection) -> Result<UIElement> {
                let pattern: UICustomNavigationPattern = self.as_ref().get_pattern()?;
                pattern.navigate(direction)
            }
        }        
    };
    gen.into()
}

pub(crate) fn impl_grid_item(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl GridItem for #name {
            fn get_containing_grid(&self) -> Result<UIElement> {
                let pattern: UIGridItemPattern = self.as_ref().get_pattern()?;
                pattern.get_containing_grid()
            }
        
            fn get_row(&self) -> Result<i32> {
                let pattern: UIGridItemPattern = self.as_ref().get_pattern()?;
                pattern.get_row()
            }
        
            fn get_column(&self) -> Result<i32> {
                let pattern: UIGridItemPattern = self.as_ref().get_pattern()?;
                pattern.get_column()
            }
        
            fn get_row_span(&self) -> Result<i32> {
                let pattern: UIGridItemPattern = self.as_ref().get_pattern()?;
                pattern.get_row_span()
            }
        
            fn get_column_span(&self) -> Result<i32> {
                let pattern: UIGridItemPattern = self.as_ref().get_pattern()?;
                pattern.get_column_span()
            }
        }        
    };
    gen.into()
}

pub(crate) fn impl_table_item(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl TableItem for #name {
            fn get_row_header_items(&self) -> Result<Vec<UIElement>> {
                let pattern: UITableItemPattern = self.as_ref().get_pattern()?;
                pattern.get_row_header_items()
            }
        
            fn get_column_header_items(&self) -> Result<Vec<UIElement>> {
                let pattern: UITableItemPattern = self.as_ref().get_pattern()?;
                pattern.get_column_header_items()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_text(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Text for #name {
            fn get_ragne_from_point(&self, pt: super::types::Point) -> Result<UITextRange> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_ragne_from_point(pt)
            }
        
            fn get_range_from_child(&self, child: &UIElement) -> Result<UITextRange> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_range_from_child(child)
            }
        
            fn get_selection(&self) -> Result<Vec<UITextRange>> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_selection()
            }
        
            fn get_visible_ranges(&self) -> Result<Vec<UITextRange>> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_visible_ranges()
            }
        
            fn get_document_range(&self) -> Result<UITextRange> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_document_range()
            }
        
            fn get_supported_text_selection(&self) -> Result<SupportedTextSelection> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_supported_text_selection()
            }
        
            fn get_range_from_annotation(&self, annotation: &UIElement) -> Result<UITextRange> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_range_from_annotation(annotation)
            }
        
            fn get_caret_range(&self) -> Result<(bool, UITextRange)> {
                let pattern: UITextPattern = self.as_ref().get_pattern()?;
                pattern.get_caret_range()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_range_value(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl RangeValue for #name {
            fn set_value(&self, value: f64) -> Result<()> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.set_value(value)
            }
        
            fn get_value(&self) -> Result<f64> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.get_value()
            }
        
            fn is_readonly(&self) -> Result<bool> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.is_readonly()
            }
        
            fn get_maximum(&self) -> Result<f64> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.get_maximum()
            }
        
            fn get_minimum(&self) -> Result<f64> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.get_minimum()
            }
        
            fn get_large_change(&self) -> Result<f64> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.get_large_change()
            }
        
            fn get_small_change(&self) -> Result<f64> {
                let pattern: UIRangeValuePattern = self.as_ref().get_pattern()?;
                pattern.get_small_change()
            }
        }
    };
    gen.into()
}

pub(crate) fn impl_dock(ast: &syn::DeriveInput) -> TokenStream {
    let name = &ast.ident;
    let gen = quote! {
        impl Dock for #name {
            fn get_dock_position(&self) -> Result<DockPosition> {
                let pattern: UIDockPattern = self.as_ref().get_pattern()?;
                pattern.get_dock_position()
            }
        
            fn set_dock_position(&self, position: DockPosition) -> Result<()> {
                let pattern: UIDockPattern = self.as_ref().get_pattern()?;
                pattern.set_dock_position(position)
            }
        }
    };
    gen.into()
}