# GitHub Setup Instructions for Excel MCP Server

This document provides step-by-step instructions to set up your local project with Git and push it to your GitHub repository.

## Prerequisites

-   Git installed on your system.
-   A GitHub account and a repository created (e.g., `https://github.com/charles69729798/mcp-excel`).

## Instructions

Ensure you are in your project directory (`C:\InsuranceProject`) in your terminal (e.g., PowerShell or Git Bash).

1.  **Initialize a Git repository (if not already done):**
    ```powershell
    git init
    ```

2.  **Configure your Git user name and email (if you haven't already):**
    Git requires your identity for every commit. Replace the placeholders with your actual email and name.
    ```powershell
    git config --global user.email "your_email@example.com"
    git config --global user.name "Your Name"
    ```

3.  **Add all project files to the staging area:**
    This command adds all new and modified files in your project directory to the staging area, preparing them for the next commit.
    ```powershell
    git add .
    ```

4.  **Commit the changes:**
    This creates a snapshot of your project's current state. The message describes the changes made.
    ```powershell
    git commit -m "Initial commit: Excel MCP Server with features and documentation"
    ```

5.  **Add the GitHub repository as a remote:**
    This links your local Git repository to your GitHub repository. Replace the URL with your actual repository URL if it's different.
    ```powershell
    git remote add origin https://github.com/charles69729798/mcp-excel
    ```
    *If you already have a remote named `origin`, you might need to use `git remote set-url origin <repository_url>` or `git remote rename origin upstream` and then add a new remote.*

6.  **Pull changes from the remote repository, allowing unrelated histories:**
    This step is crucial if your GitHub repository was initialized with a `README.md` or other files, as it brings those changes into your local repository. The `--allow-unrelated-histories` flag is used because your local and remote repositories started with separate histories.
    ```powershell
    git pull origin main --allow-unrelated-histories
    ```
    *   **If a merge conflict occurs (e.g., in `README.md`):**
        *   Open the conflicted file (e.g., `README.md`) in a text editor.
        *   You will see conflict markers (`<<<<<<<`, `=======`, `>>>>>>>`). Manually resolve the conflict by choosing the content you want to keep. For `README.md`, it's recommended to keep the more detailed content provided by the agent.
        *   Save the file.
        *   Add the resolved file to the staging area: `git add README.md`
        *   Commit the merge: `git commit` (Git will provide a default merge message, which you can accept by saving and closing the editor).

7.  **Push the changes to GitHub:**
    This uploads your local commits to your GitHub repository. The `-u` flag sets the upstream branch, so future `git push` and `git pull` commands will work without specifying `origin main`.
    ```powershell
    git push -u origin master:main
    ```
    *(Note: `main` is the common default branch name. If your GitHub repository uses `master` as its default branch, you might need to use `git push -u origin master` or `git push -u origin master:master`.)*

## Troubleshooting

-   **`fatal: refusing to merge unrelated histories`:** This is handled by step 6. Ensure you use the `--allow-unrelated-histories` flag.
-   **`error: src refspec main does not match any`:** This usually means your local branch is not named `main` (it might be `master`). Adjust the push command accordingly (e.g., `git push -u origin master:main`).
-   **`Author identity unknown`:** This means you need to configure your Git user name and email (Step 2).
-   **`! [rejected] ... (fetch first)`:** This means the remote repository has changes you don't have locally. Use `git pull` (Step 6) to integrate them before pushing.