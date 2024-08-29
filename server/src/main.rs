/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Main module

// #![forbid(unsafe_code)]
#![warn(clippy::pedantic)]
#![allow(clippy::module_name_repetitions)]

use std::net::SocketAddr;
use std::pin::pin;
use std::str::FromStr;
use std::sync::Arc;

use axum::extract::DefaultBodyLimit;
use axum::routing::{get, post};
use axum::Router;
use log::info;

use crate::application::Application;
use crate::routes::AxumState;
use crate::utils::sigterm::waiting_sigterm;

mod application;
mod model;
mod routes;
mod utils;
mod webapp;

/// The name of this program
pub const CRATE_NAME: &str = env!("CARGO_PKG_NAME");
/// The commit that was used to build the application
pub const GIT_HASH: &str = env!("GIT_HASH");
/// The git tag tag that was used to build the application
pub const GIT_TAG: &str = env!("GIT_TAG");

/// Main payload for serving the application
async fn main_serve_app(application: Arc<Application>) -> Result<(), std::io::Error> {
    // web application
    let webapp_resources = webapp::get_resources();
    let body_limit = application.configuration.web_body_limit;
    let socket_addr = SocketAddr::new(
        application.configuration.web_listenon_ip,
        application.configuration.web_listenon_port,
    );
    let state = Arc::new(AxumState {
        application,
        webapp_resources,
    });
    let app = Router::new()
        // api version
        .route("/version", get(crate::routes::get_version))
        // API
        .nest(
            "/api/v1",
            Router::new()
                .route("/tasks", post(crate::routes::launch_task))
                .route("/tasks/:task_id/updates", get(crate::routes::observe_task))
                .route("/tasks/:task_id/download", get(crate::routes::download_task_result)),
        )
        // fall back to serving the index
        .fallback(crate::routes::get_webapp_resource)
        .layer(DefaultBodyLimit::max(body_limit))
        .with_state(state);
    axum::serve(
        tokio::net::TcpListener::bind(socket_addr)
            .await
            .unwrap_or_else(|_| panic!("failed to bind {socket_addr}")),
        app.into_make_service_with_connect_info::<SocketAddr>(),
    )
    .await
}

fn setup_log() {
    let log_date_time_format =
        std::env::var("REGISTRY_LOG_DATE_TIME_FORMAT").unwrap_or_else(|_| String::from("[%Y-%m-%d %H:%M:%S]"));
    let log_level = std::env::var("REGISTRY_LOG_LEVEL")
        .map(|v| log::LevelFilter::from_str(&v).expect("invalid REGISTRY_LOG_LEVEL"))
        .unwrap_or(log::LevelFilter::Info);
    fern::Dispatch::new()
        .filter(move |metdata| {
            let target = metdata.target();
            target.starts_with("cratery") || target.starts_with("cenotelie")
        })
        .format(move |out, message, record| {
            out.finish(format_args!(
                "{}\t{}\t{}",
                chrono::Local::now().format(&log_date_time_format),
                record.level(),
                message
            ));
        })
        .level(log_level)
        .chain(std::io::stdout())
        .apply()
        .expect("log configuration failed");
}

/// Main entry point
#[tokio::main]
async fn main() {
    setup_log();
    info!("{} commit={} tag={}", CRATE_NAME, GIT_HASH, GIT_TAG);

    let application = crate::application::Application::launch();

    let server = pin!(main_serve_app(application));

    let _ = waiting_sigterm(server).await;
}
