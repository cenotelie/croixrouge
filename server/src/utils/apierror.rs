/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Definition of the error type for API requests

use std::backtrace::Backtrace;
use std::fmt::{Display, Formatter};

use serde_derive::{Deserialize, Serialize};

/// Describes an API error
#[derive(Serialize, Deserialize, Debug)]
pub struct ApiError {
    /// The associated HTTP error code
    pub http: u16,
    /// A custom error message
    pub message: String,
    /// Optional details for the error
    pub details: Option<String>,
    /// The backtrace when the error was produced
    #[serde(skip_serializing, skip_deserializing)]
    pub backtrace: Option<Backtrace>,
}

impl ApiError {
    /// Creates a new error
    #[allow(clippy::needless_pass_by_value)]
    #[must_use]
    pub fn new<M: ToString>(http: u16, message: M, details: Option<String>) -> Self {
        Self {
            http,
            message: message.to_string(),
            details,
            backtrace: Some(Backtrace::capture()),
        }
    }
}

impl Display for ApiError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let details = self.details.as_ref().map_or("", std::convert::AsRef::as_ref);
        write!(f, "{} ({})", &self.message, &details)
    }
}

impl Clone for ApiError {
    fn clone(&self) -> Self {
        Self {
            http: self.http,
            message: self.message.clone(),
            details: self.details.clone(),
            backtrace: None,
        }
    }
}

impl<E> From<E> for ApiError
where
    E: std::error::Error,
{
    fn from(err: E) -> Self {
        Self::new(500, "The operation failed in the backend.", Some(err.to_string()))
    }
}

/// Error when the requested user cannot be found
#[must_use]
pub fn error_not_found() -> ApiError {
    ApiError::new(404, "The requested resource cannot be found.", None)
}
