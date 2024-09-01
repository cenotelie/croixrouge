/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Data types for tasks

use std::path::PathBuf;
use std::process::Stdio;

use chrono::{Local, NaiveDateTime};
use serde_derive::{Deserialize, Serialize};
use tokio::process::Command;

use crate::utils::apierror::{error_backend_failure, specialize, ApiError};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TaskInputs {
    #[serde(rename = "fileName")]
    pub file_name: String,
    #[serde(rename = "fileContent")]
    pub file_content: Vec<u8>,
    pub mode: i32,
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TaskStatus {
    #[default]
    Created,
    Executing,
    Completed,
    Failed,
}

impl TaskStatus {
    pub fn is_final(self) -> bool {
        matches!(self, TaskStatus::Completed | TaskStatus::Failed)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Task {
    pub identifier: String,
    pub status: TaskStatus,
    #[serde(rename = "lastUpdate")]
    pub last_update: NaiveDateTime,
    pub inputs: TaskInputs,
    pub output: String,
}

impl Task {
    pub async fn new(inputs: TaskInputs) -> Result<Self, ApiError> {
        let identifier = uuid::Uuid::new_v4().to_string();
        let mut path = PathBuf::from("/home/croixrouge");
        path.push("tasks");
        path.push(&identifier);
        tokio::fs::create_dir_all(&path).await?;
        path.push("payload.xlsx");
        tokio::fs::write(&path, &inputs.file_content).await?;

        Ok(Self {
            identifier,
            status: TaskStatus::default(),
            last_update: Local::now().naive_utc(),
            inputs: TaskInputs {
                file_content: Vec::new(),
                ..inputs
            },
            output: String::new(),
        })
    }

    pub fn touch(&mut self) {
        self.last_update = Local::now().naive_utc();
    }

    pub async fn load_result(&self) -> Result<Vec<u8>, ApiError> {
        let mut path = PathBuf::from("/home/croixrouge");
        path.push("tasks");
        path.push(&self.identifier);
        path.push("payload.xlsx");
        let content = tokio::fs::read(&path).await?;
        Ok(content)
    }

    /// Executes this task
    pub async fn execute(&self) -> Result<(), ApiError> {
        let mut path = PathBuf::from("/home/croixrouge");
        path.push("tasks");
        path.push(&self.identifier);
        let command = Command::new("/home/croixrouge/payload/DistributionCR")
            .current_dir(path)
            .args(["payload.xlsx", &self.inputs.mode.to_string()])
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .output()
            .await?;
        if command.status.success() {
            Ok(())
        } else {
            let mut details = String::from_utf8(command.stdout)?;
            details.push_str(&String::from_utf8(command.stderr)?);
            Err(specialize(error_backend_failure(), details))
        }
    }

    /// Removes all stored data
    pub async fn delete(&self) -> Result<(), ApiError> {
        let mut path = PathBuf::from("/home/croixrouge");
        path.push("tasks");
        path.push(&self.identifier);
        tokio::fs::remove_dir_all(path).await?;
        Ok(())
    }
}
