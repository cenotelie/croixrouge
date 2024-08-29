/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Data model

pub mod config;
pub mod errors;
pub mod tasks;

use serde_derive::{Deserialize, Serialize};

/// The object representing the application version
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct AppVersion {
    /// The changeset that was used to build the app
    pub commit: String,
    /// The version tag, if any
    pub tag: String,
}
