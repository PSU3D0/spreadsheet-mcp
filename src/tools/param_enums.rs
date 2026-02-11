use schemars::JsonSchema;
use serde::de;
use serde::{Deserialize, Serialize};
use std::fmt;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
#[derive(Default)]
pub enum BatchMode {
    #[default]
    Apply,
    Preview,
}

impl BatchMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Apply => "apply",
            Self::Preview => "preview",
        }
    }

    pub fn is_preview(self) -> bool {
        matches!(self, Self::Preview)
    }
}

impl fmt::Display for BatchMode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl<'de> Deserialize<'de> for BatchMode {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: de::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.to_ascii_lowercase().as_str() {
            "apply" => Ok(Self::Apply),
            "preview" => Ok(Self::Preview),
            other => Err(de::Error::unknown_variant(other, &["apply", "preview"])),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
#[derive(Default)]
pub enum ReplaceMatchMode {
    #[default]
    Exact,
    Contains,
}

impl ReplaceMatchMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Exact => "exact",
            Self::Contains => "contains",
        }
    }
}

impl<'de> Deserialize<'de> for ReplaceMatchMode {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: de::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.to_ascii_lowercase().as_str() {
            "exact" => Ok(Self::Exact),
            "contains" => Ok(Self::Contains),
            other => Err(de::Error::unknown_variant(other, &["exact", "contains"])),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
#[derive(Default)]
pub enum FillDirection {
    Down,
    Right,
    #[default]
    Both,
}

impl FillDirection {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Down => "down",
            Self::Right => "right",
            Self::Both => "both",
        }
    }
}

impl<'de> Deserialize<'de> for FillDirection {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: de::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.to_ascii_lowercase().as_str() {
            "down" => Ok(Self::Down),
            "right" => Ok(Self::Right),
            "both" => Ok(Self::Both),
            other => Err(de::Error::unknown_variant(
                other,
                &["down", "right", "both"],
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
#[derive(Default)]
pub enum FormulaRelativeMode {
    #[default]
    Excel,
    AbsCols,
    AbsRows,
}

impl<'de> Deserialize<'de> for FormulaRelativeMode {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: de::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.to_ascii_lowercase().as_str() {
            "excel" => Ok(Self::Excel),
            "abs_cols" | "abscols" | "columns_absolute" => Ok(Self::AbsCols),
            "abs_rows" | "absrows" | "rows_absolute" => Ok(Self::AbsRows),
            other => Err(de::Error::unknown_variant(
                other,
                &["excel", "abs_cols", "abs_rows"],
            )),
        }
    }
}

impl From<FormulaRelativeMode> for crate::formula::pattern::RelativeMode {
    fn from(value: FormulaRelativeMode) -> Self {
        match value {
            FormulaRelativeMode::Excel => Self::Excel,
            FormulaRelativeMode::AbsCols => Self::AbsCols,
            FormulaRelativeMode::AbsRows => Self::AbsRows,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, JsonSchema)]
#[serde(rename_all = "snake_case")]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

impl<'de> Deserialize<'de> for PageOrientation {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: de::Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        match s.to_ascii_lowercase().as_str() {
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            other => Err(de::Error::unknown_variant(
                other,
                &["portrait", "landscape"],
            )),
        }
    }
}

impl PageOrientation {
    pub fn to_umya(self) -> umya_spreadsheet::OrientationValues {
        match self {
            Self::Portrait => umya_spreadsheet::OrientationValues::Portrait,
            Self::Landscape => umya_spreadsheet::OrientationValues::Landscape,
        }
    }
}
