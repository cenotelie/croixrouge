/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

//! Main application

use std::sync::{Arc, Mutex};
use std::time::Duration;

use log::error;
use tokio::sync::mpsc::{Receiver, Sender};

use crate::model::config::Configuration;
use crate::model::tasks::{Task, TaskInputs, TaskStatus};
use crate::utils::apierror::{error_not_found, ApiError};

/// The state of this application for axum
pub struct Application {
    /// The configuration
    pub configuration: Arc<Configuration>,
    /// The ongoing tasks
    pub tasks: Mutex<Vec<TaskData>>,
    /// The sender to send tasks to the worker
    pub task_sender: Sender<String>,
}

pub struct TaskData {
    pub task: Task,
    pub observer: Option<Sender<TaskStatus>>,
}

const TASK_WORKER_CHANNEL_SIZE: usize = 10;

impl Application {
    /// Creates a new application
    pub fn launch() -> Arc<Self> {
        // load configuration
        let configuration = Arc::new(Configuration::from_env());
        let (task_sender, task_receiver) = tokio::sync::mpsc::channel(TASK_WORKER_CHANNEL_SIZE);
        let app = Arc::new(Self {
            configuration,
            tasks: Mutex::new(Vec::new()),
            task_sender,
        });
        tokio::spawn({
            let app = app.clone();
            Box::pin(async move {
                tasks_worker(app, task_receiver).await;
            })
        });
        app
    }

    /// Creates and queues a new task
    pub async fn create_task(&self, inputs: TaskInputs) -> Result<Task, ApiError> {
        let task = Task::new(inputs).await?;
        self.tasks.lock().unwrap().push(TaskData {
            task: task.clone(),
            observer: None,
        });
        self.task_sender.send(task.identifier.clone()).await?;
        Ok(task)
    }

    /// Starts observing the progress of a task
    pub async fn observe_task(&self, task_id: &str) -> Result<Receiver<TaskStatus>, ApiError> {
        let (receiver, update) = {
            let mut lock = self.tasks.lock().unwrap();
            let data = lock
                .iter_mut()
                .find(|data| data.task.identifier == task_id)
                .ok_or_else(error_not_found)?;

            let (sender, receiver) = tokio::sync::mpsc::channel(4);
            if data.task.status.is_final() {
                (receiver, Some((sender, data.task.status)))
            } else {
                data.observer = Some(sender.clone());
                (receiver, None)
            }
        };
        if let Some((sender, status)) = update {
            sender.send(status).await?;
        }
        Ok(receiver)
    }

    /// Update a task's status
    async fn on_task_update(&self, task_id: &str, status: TaskStatus) -> Result<(), ApiError> {
        let sender = self
            .tasks
            .lock()
            .unwrap()
            .iter_mut()
            .find(|d| d.task.identifier == task_id)
            .and_then(|data| {
                data.task.status = status;
                data.task.touch();
                if status.is_final() {
                    data.observer.take()
                } else {
                    data.observer.clone()
                }
            });
        if let Some(sender) = sender {
            sender.send(status).await?;
        }
        Ok(())
    }

    /// Gets the result of task
    pub async fn get_task_result(&self, task_id: &str) -> Result<(String, Vec<u8>), ApiError> {
        let task = self
            .tasks
            .lock()
            .unwrap()
            .iter()
            .find(|d| d.task.identifier == task_id)
            .ok_or_else(error_not_found)?
            .task
            .clone();
        let data = task.load_result().await?;
        Ok((task.inputs.file_name, data))
    }
}

/// The task worker
async fn tasks_worker(app: Arc<Application>, mut task_receiver: Receiver<String>) {
    while let Some(task_id) = task_receiver.recv().await {
        let status = match tasks_worker_task(&app, &task_id).await {
            Ok(()) => TaskStatus::Completed,
            Err(e) => {
                error!("{e}");
                if let Some(backtrace) = &e.backtrace {
                    error!("{backtrace}");
                }
                TaskStatus::Failed
            }
        };
        // set final status
        let _ = app.on_task_update(&task_id, status).await;
    }
}

async fn tasks_worker_task(app: &Application, task_id: &str) -> Result<(), ApiError> {
    let _task = app
        .tasks
        .lock()
        .unwrap()
        .iter()
        .find_map(|d| {
            if d.task.identifier == task_id {
                Some(d.task.clone())
            } else {
                None
            }
        })
        .ok_or_else(error_not_found)?;
    app.on_task_update(task_id, TaskStatus::Executing).await?;
    tokio::time::sleep(Duration::from_secs(5)).await;
    Ok(())
}
