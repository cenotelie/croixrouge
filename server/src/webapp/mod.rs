/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Encapsulation of the web application files

use crate::utils::axum::embedded::{get_content_type, Resource, Resources};

macro_rules! add {
    ($resources: expr, $name: literal) => {
        $resources.data.insert(
            $name,
            Resource {
                file_name: $name,
                content_type: get_content_type($name),
                content: include_bytes!($name),
            },
        );
    };
}

/// Gets the resources to serve for the web application
pub fn get_resources() -> Resources {
    let mut resources = Resources::with_fallback("index.html");
    // HTML
    add!(resources, "index.html");
    add!(resources, "index.js");
    add!(resources, "index.css");
    // images
    add!(resources, "cenotelie.png");
    add!(resources, "favicon.png");
    resources
}
