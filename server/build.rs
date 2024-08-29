/*******************************************************************************
 * Copyright (c) 2024 Cénotélie (cenotelie.fr)
 ******************************************************************************/

use std::process::Command;

fn main() {
    if let Ok(output) = Command::new("git").args(["rev-parse", "HEAD"]).output() {
        let value = String::from_utf8(output.stdout).unwrap();
        println!("cargo:rustc-env=GIT_HASH={value}");
    }
    if let Ok(output) = Command::new("git").args(["tag", "-l", "--points-at", "HEAD"]).output() {
        let value = String::from_utf8(output.stdout).unwrap();
        println!("cargo:rustc-env=GIT_TAG={value}");
    }
}
