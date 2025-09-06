use clap::ValueEnum;
use serde::{Deserialize, Serialize};

/// Input parameters (all in inches)
#[derive(Debug, Clone, Copy, Serialize, Deserialize, ValueEnum)]
pub enum UnitSystem {
    Inch,
    Cm,
}

impl UnitSystem {
    pub fn as_str(&self) -> &'static str {
        match self {
            UnitSystem::Inch => "in",
            UnitSystem::Cm => "cm",
        }
    }
}

#[derive(Debug)]
pub struct BookParams {
    pub width: f64,
    pub height: f64,
    pub unit_system: UnitSystem,
    pub pages: i64,
}

#[derive(Debug)]
pub struct BookBindingConstant {
    /// per-edge bleed (usually 0.125")
    pub bleed_cover: f64,
    /// cover margin (applied equally to top/bottom/left/right)
    pub margin_cover: f64,
    /// thickness per page
    pub thickness: f64,
    /// inner margin (gutter)
    pub gutter: f64,
    /// outer margin
    pub margin_inner: f64,
}

#[derive(Debug)]
pub struct Book {
    pub params: BookParams,
    pub binding: BookBindingConstant,
}

#[derive(Debug)]
pub struct Size {
    pub width: f64,
    pub height: f64,
}

#[derive(Debug)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

impl Book {
    pub fn new(params: BookParams, binding: BookBindingConstant) -> Self {
        Self { params, binding }
    }

    /// Get spine width
    pub fn get_spine_width(&self) -> f64 {
        self.params.pages as f64 * self.binding.thickness
    }

    /// Get cover size
    pub fn get_cover_size(&self) -> Size {
        let spine = self.get_spine_width();
        let w = 2.0 * self.params.width
            + 2.0 * self.binding.bleed_cover
            + 2.0 * self.binding.margin_cover
            + spine;
        let h = self.params.height
            + 2.0 * self.binding.bleed_cover
            + 2.0 * self.binding.margin_cover;
        Size { width: w, height: h }
    }

    /// Get safe area size
    pub fn get_safe_area_size(&self) -> Size {
        let w = self.params.width - (self.binding.gutter + self.binding.margin_inner);
        let h = self.params.height - (2.0 * self.binding.margin_inner);
        Size { width: w, height: h }
    }

    /// Get safe area rect (if is_left: true = left page (verso), false = right page (recto))
    pub fn get_safe_area(&self, is_left: bool) -> Rect {
        let safe = self.get_safe_area_size();
        let x = if is_left { self.binding.margin_inner } else { self.binding.gutter };
        let y = self.binding.margin_inner;

        Rect {
            x,
            y,
            width: safe.width,
            height: safe.height,
        }
    }
}

// const THICKNESS_PREMIUM: f64 = 0.002347;
const THICKNESS_WHITE: f64 = 0.002252;
const THICKNESS_CREAM: f64 = 0.0025;

const BINDING_PARAMS_KDP_WHITE: BookBindingConstant = BookBindingConstant {
    bleed_cover: 0.125,         // KDP default
    margin_cover: 0.125,        // conservative cover margin when bleed is present
    thickness: THICKNESS_WHITE, // example: 120p B/W White (0.002252 * 120 ≈ 0.270; varies by vendor)
    gutter: 0.375,              // inner margin
    margin_inner: 0.25,         // outer margin safety margin
};

const BINDING_PARAMS_KDP_CREAM: BookBindingConstant = BookBindingConstant {
    bleed_cover: 0.125,         // KDP default
    margin_cover: 0.125,        // conservative cover margin when bleed is present
    thickness: THICKNESS_CREAM, // example: 120p B/W Cream (0.0025 * 120 ≈ 0.300; varies by vendor)
    gutter: 0.375,              // inner margin
    margin_inner: 0.25,         // outer margin safety margin
};
