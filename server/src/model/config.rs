/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Module for configuration management

use std::net::{IpAddr, Ipv4Addr};
use std::str::FromStr;

use serde_derive::{Deserialize, Serialize};

use crate::model::errors::MissingEnvVar;

/// Gets the value for an environment variable
pub fn get_var<T: AsRef<str>>(name: T) -> Result<String, MissingEnvVar> {
    let key = name.as_ref();
    std::env::var(key).map_err(|original| MissingEnvVar {
        original,
        var_name: key.to_string(),
    })
}

/// A configuration for the registry
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct Configuration {
    /// The log level to use
    #[serde(rename = "logLevel")]
    pub log_level: String,
    /// The datetime format to use when logging
    #[serde(rename = "logDatetimeFormat")]
    pub log_datetime_format: String,
    /// The IP to bind for the web server
    #[serde(rename = "webListenOnIp")]
    pub web_listenon_ip: IpAddr,
    /// The port to bind for the web server
    #[serde(rename = "webListenOnPort")]
    pub web_listenon_port: u16,
    /// The maximum size for the body of incoming requests
    #[serde(rename = "webBodyLimit")]
    pub web_body_limit: usize,
    /// The directory to use for hot reloading the web application
    #[serde(rename = "webHotReload")]
    pub web_hot_reload: Option<String>,
}

impl Configuration {
    /// Gets the configuration from environment variables
    ///
    /// # Errors
    ///
    /// Return a `VarError` when an expected environment variable is not present
    pub fn from_env() -> Self {
        Self {
            log_level: get_var("CROIXROUGE_LOG_LEVEL").unwrap_or_else(|_| String::from("INFO")),
            log_datetime_format: get_var("CROIXROUGE_LOG_DATE_TIME_FORMAT")
                .unwrap_or_else(|_| String::from("[%Y-%m-%d %H:%M:%S]")),
            web_listenon_ip: get_var("CROIXROUGE_WEB_LISTENON_IP").map_or_else(
                |_| IpAddr::V4(Ipv4Addr::UNSPECIFIED),
                |s| IpAddr::from_str(&s).expect("invalud CROIXROUGE_WEB_LISTENON_IP"),
            ),
            web_listenon_port: get_var("CROIXROUGE_WEB_LISTENON_PORT")
                .map(|s| s.parse().expect("invalid CROIXROUGE_WEB_LISTENON_PORT"))
                .unwrap_or(80),
            web_body_limit: get_var("CROIXROUGE_WEB_BODY_LIMIT")
                .map(|s| s.parse().expect("invalid CROIXROUGE_WEB_BODY_LIMIT"))
                .unwrap_or(10 * 1024 * 1024),
            web_hot_reload: get_var("CROIXROUGE_WEB_HOT_RELOAD").ok(),
        }
    }
}
