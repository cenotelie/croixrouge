/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Implementation of axum routes to expose the application

use std::borrow::Cow;
use std::path::PathBuf;
use std::sync::Arc;

use axum::body::{Body, Bytes};
use axum::extract::{Path, Query, State};
use axum::http::header::HeaderName;
use axum::http::{header, HeaderValue, Request, StatusCode};
use axum::response::{IntoResponse, Response};
use axum::Json;
use futures::StreamExt;
use serde_derive::Deserialize;
use tokio_stream::wrappers::ReceiverStream;

use crate::application::Application;
use crate::model::tasks::{Task, TaskInputs};
use crate::model::AppVersion;
use crate::utils::apierror::ApiError;
use crate::utils::axum::embedded::{get_content_type, Resource, Resources};
use crate::utils::axum::sse::{Event, ServerSentEventStream};
use crate::utils::axum::{response, response_error, ApiResult};

/// The state of this application for axum
pub struct AxumState {
    /// The main application
    pub application: Arc<Application>,
    /// The static resources for the web app
    pub webapp_resources: Resources,
}

enum WebappResource {
    Embedded(Resource),
    HotReload(String, Vec<u8>),
}

impl AxumState {
    /// Gets the resource in the web app for the specified path
    async fn get_webapp_resource(&self, path: &str) -> Option<WebappResource> {
        if let Some(hot_reload_path) = self.application.configuration.web_hot_reload.as_ref() {
            let mut final_path = PathBuf::from(hot_reload_path);
            for element in path.split('/') {
                final_path.push(element);
            }
            let file_name = final_path.file_name().and_then(|n| n.to_str()).unwrap();
            let content_type = get_content_type(file_name);
            let data = tokio::fs::read(&final_path).await.ok()?;
            Some(WebappResource::HotReload(content_type.to_string(), data))
        } else {
            let resource = self.webapp_resources.get(path).cloned()?;
            Some(WebappResource::Embedded(resource))
        }
    }
}

/// Gets the favicon
pub async fn get_webapp_resource(
    State(state): State<Arc<AxumState>>,
    request: Request<Body>,
) -> Result<(StatusCode, [(HeaderName, HeaderValue); 2], Cow<'static, [u8]>), StatusCode> {
    let path = &request.uri().path()[1..];
    match state.get_webapp_resource(path).await {
        Some(WebappResource::Embedded(resource)) => Ok((
            StatusCode::OK,
            [
                (header::CONTENT_TYPE, HeaderValue::from_static(resource.content_type)),
                (header::CACHE_CONTROL, HeaderValue::from_static("max-age=3600")),
            ],
            Cow::Borrowed(resource.content),
        )),
        Some(WebappResource::HotReload(content_type, content)) => Ok((
            StatusCode::OK,
            [
                (header::CONTENT_TYPE, HeaderValue::from_str(&content_type).unwrap()),
                (header::CACHE_CONTROL, HeaderValue::from_static("max-age=3600")),
            ],
            Cow::Owned(content),
        )),
        None => Err(StatusCode::NOT_FOUND),
    }
}

/// Gets the version data for the application
///
/// # Errors
///
/// Always return the `Ok` variant, but use `Result` for possible future usage.
pub async fn get_version() -> ApiResult<AppVersion> {
    response(Ok(AppVersion {
        commit: crate::GIT_HASH.to_string(),
        tag: crate::GIT_TAG.to_string(),
    }))
}

#[derive(Deserialize)]
pub struct LaunchTaskQuery {
    name: String,
    mode: i32,
}

pub async fn launch_task(
    State(state): State<Arc<AxumState>>,
    Query(LaunchTaskQuery { name, mode }): Query<LaunchTaskQuery>,
    body: Bytes,
) -> ApiResult<Task> {
    response(
        state
            .application
            .create_task(TaskInputs {
                file_name: name,
                mode,
                file_content: body.to_vec(),
            })
            .await,
    )
}

pub async fn observe_task(
    State(state): State<Arc<AxumState>>,
    Path(task_id): Path<String>,
) -> Result<Response, (StatusCode, Json<ApiError>)> {
    let receiver = match state.application.observe_task(&task_id).await {
        Ok(r) => r,
        Err(e) => return Err(response_error(e)),
    };
    let stream = ServerSentEventStream::new(ReceiverStream::new(receiver).map(Event::from_data));
    Ok(stream.into_response())
}

pub async fn download_task_result(
    State(state): State<Arc<AxumState>>,
    Path(task_id): Path<String>,
) -> Result<(StatusCode, [(HeaderName, HeaderValue); 2], Vec<u8>), (StatusCode, Json<ApiError>)> {
    match state.application.get_task_result(&task_id).await {
        Ok((file_name, data)) => Ok((
            StatusCode::OK,
            [
                (
                    axum::http::header::CONTENT_TYPE,
                    HeaderValue::from_static("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ),
                (
                    axum::http::header::CONTENT_DISPOSITION,
                    HeaderValue::from_str(&format!("attachment; filename=\"{file_name}\"")).unwrap(),
                ),
            ],
            data,
        )),
        Err(e) => Err(response_error(e)),
    }
}
