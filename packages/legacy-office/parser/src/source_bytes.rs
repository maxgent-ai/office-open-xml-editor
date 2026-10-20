//! Source acquisition is fallible, but file length is not admission policy.
use wasm_bindgen::prelude::*;

pub(crate) fn copy(data: js_sys::Uint8Array) -> Result<Vec<u8>, JsValue> {
    let mut bytes = allocate(data.length() as usize).map_err(|error| JsValue::from_str(&error))?;
    data.copy_to(&mut bytes);
    Ok(bytes)
}

fn allocate(length: usize) -> Result<Vec<u8>, String> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|_| crate::rejection::resource("legacy source allocation failed"))?;
    bytes.resize(length, 0);
    Ok(bytes)
}

#[cfg(test)]
mod tests {
    #[test]
    fn source_copy_allocation_failure_is_classified() {
        assert_eq!(
            super::allocate(usize::MAX).unwrap_err(),
            "LEGACY_RESOURCE_LIMIT:legacy source allocation failed"
        );
        assert_eq!(super::allocate(3).unwrap(), [0, 0, 0]);
    }
}
