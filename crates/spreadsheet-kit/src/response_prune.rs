use anyhow::Result;
use schemars::{JsonSchema, Schema, SchemaGenerator};
use serde::Serialize;
use serde::ser::Error as _;
use serde_json::Value;
use std::borrow::Cow;

pub fn to_pruned_value<T: Serialize>(value: &T) -> Result<Value> {
    let mut json = serde_json::to_value(value)?;
    prune_non_structural_empties(&mut json);
    Ok(json)
}

#[derive(Debug, Clone)]
pub struct Pruned<T>(pub T);

impl<T: Serialize> Serialize for Pruned<T> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        let mut value = serde_json::to_value(&self.0).map_err(S::Error::custom)?;
        prune_non_structural_empties(&mut value);
        value.serialize(serializer)
    }
}

impl<T: JsonSchema> JsonSchema for Pruned<T> {
    fn schema_name() -> Cow<'static, str> {
        T::schema_name()
    }

    fn json_schema(generator: &mut SchemaGenerator) -> Schema {
        T::json_schema(generator)
    }
}

pub fn prune_non_structural_empties(value: &mut Value) {
    match value {
        Value::Object(map) => {
            let mut remove_keys = Vec::new();
            for (key, child) in map.iter_mut() {
                prune_non_structural_empties(child);
                if child.is_null() {
                    remove_keys.push(key.clone());
                    continue;
                }
                if let Value::Object(obj) = child
                    && obj.is_empty()
                {
                    remove_keys.push(key.clone());
                    continue;
                }
                if let Value::Array(arr) = child
                    && arr.is_empty()
                    && key != "warnings"
                {
                    remove_keys.push(key.clone());
                }
            }
            for key in remove_keys {
                map.remove(&key);
            }
        }
        Value::Array(items) => {
            for child in items.iter_mut() {
                prune_non_structural_empties(child);
            }
        }
        _ => {}
    }
}
