//! Experimental direct `.ppt` presentation-source boundary.
//!
//! This API does not create an OOXML package. It owns a bounded native cursor that projects the supported passive PPT
//! subset directly into the existing PPTX renderer models. The caller retains
//! ownership of the original JavaScript bytes; the fallible constructor copy is
//! released after construction once the session owns its admitted streams.

use crate::cfb::CompoundFile;
use crate::ppt::{direct_cursor::DirectCursor, direct_session::DirectSession};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct LegacyPptPresentation {
    cursor: DirectCursor,
}

#[wasm_bindgen]
impl LegacyPptPresentation {
    #[wasm_bindgen(constructor)]
    pub fn new(data: js_sys::Uint8Array) -> Result<LegacyPptPresentation, JsValue> {
        console_error_panic_hook::set_once();
        let source = crate::source_bytes::copy(data)?;
        let cfb = CompoundFile::open(&source)
            .map_err(|error| js_error(&crate::rejection::unsupported(error)))?;
        let session = DirectSession::new(&cfb).map_err(|error| js_error(&error))?;
        Ok(Self {
            cursor: DirectCursor::new(session),
        })
    }

    pub fn presentation_bootstrap(&self) -> Result<Vec<u8>, JsValue> {
        self.cursor.presentation_bootstrap().map_err(string_error)
    }

    pub fn pull_slide(
        &mut self,
        slide_index: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: u32,
    ) -> Result<Vec<u8>, JsValue> {
        self.cursor
            .pull_slide(
                slide_index as usize,
                operation_id,
                generation,
                byte_credit as usize,
            )
            .map_err(string_error)
    }

    pub fn acknowledge_slide(&mut self, operation_id: u32, generation: u32) -> Result<(), JsValue> {
        self.cursor
            .acknowledge_slide(operation_id, generation)
            .map_err(string_error)
    }

    pub fn cancel_slide(&mut self) -> Result<(), JsValue> {
        self.cursor.cancel_slide().map_err(string_error)
    }

    pub fn close_presentation_session(&mut self) -> Result<(), JsValue> {
        self.cursor
            .close_presentation_session()
            .map_err(string_error)
    }

    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.cursor.assert_healthy().map_err(string_error)
    }

    pub fn extract_image(&self, path: &str) -> Result<Vec<u8>, JsValue> {
        self.cursor.extract_image(path).map_err(string_error)
    }

    pub fn slide_cursor_resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        self.cursor.assert_healthy().map_err(string_error)?;
        Err(js_error("slide cursor usage is unavailable"))
    }
}

fn string_error(error: String) -> JsValue {
    js_error(&error)
}

fn js_error(error: &str) -> JsValue {
    JsValue::from_str(error)
}
