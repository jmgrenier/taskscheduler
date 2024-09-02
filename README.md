# Task Scheduler

## What is it?

This is the code I use in production for automating the creation, update and deletion of scheduled task on a Windows server during deployments.
Since there aren't a lot of ressources online for interacting with win32com, I decided to share in the hopes it could help someone.
The code is provided as-is; it should work out-of-the-box, but you're encouraged to adapt it to your needs.

## Requirements

The module only depends on `pywin32`. It has been tested with Python 3.12.* on Windows 11 and Windows Server 2012 R2.

## Usage

Tasks definition used by TaskScheduler.build are dicts. Their format should mirror the [object's model](https://learn.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-objects). Most attributes can be omitted, their default values will be identical to the Task Manager's creation wizard in most cases.

To achieve parity with XML config files, constants and bit flags can be specified using their name instead of their integer value.

### Example

Given the following example task definitions:

```yaml
- Path: \myApp\simplest_task
  RegistrationInfo:
    Description: A task only needs a path and an action to be registered!
  Actions:
    - Type: EXEC
      Path: path\to\exe

- Path: \myApp\daily_task
  RegistrationInfo:
    Description: A task that executes a Python script with a daily trigger and a 5 minute execution limit.
  Settings:
    ExecutionTimeLimit: PT5M
  Triggers:
    - Type: DAILY
      StartBoundary: 2024-09-03T00:00:00  # datetimes will be converted to iso format
  Actions:
    - Type: EXEC
      Path: path\to\python.exe
      Arguments: path\to\myApp\daily_task.py
      WorkingDirectory: path\to\myApp

- Path: \myApp\multi_task
  RegistrationInfo:
    Description: >-
      A task that executes a Python script with multiple triggers and a 1 hour execution limit.
      If it fails, it will be retried 2 times with a 1 minute interval between each try.
      It will start as soon as possible if the last execution was missed.
  Settings:
    ExecutionTimeLimit: PT1H
    RestartCount: 2
    RestartInterval: PT1M
    StartWhenAvailable: true
  Triggers:
    - Type: DAILY
      StartBoundary: 2024-09-03T01:00:00
    - Type: MONTHLY
      StartBoundary: 2024-09-04T02:00:00
      DaysOfMonth:
        - "1" # must be str
        - "2"
    - Type: TIME
      StartBoundary: 2024-09-05T03:00:00
      Repetition:
        Interval: PT10M
  Actions:
    - Type: EXEC
      Path: path\to\python.exe
      Arguments: path\to\myApp\multi_task.py
      WorkingDirectory: path\to\myApp
```

You can:

```python
# Initialize
scheduler = TaskScheduler()

# Load your task definitions
tasks_definitions = ...

# Build
tasks = [scheduler.build(task_def) for task_def in task_definitions]

# Register
for task in tasks:
    scheduler.register(task)

# Delete
scheduler.delete_task('\\path\\to\\task')  # Only the path is needed

# Sync
#   - Creates tasks in list but missing from folder
#   - Updates tasks in list already in folder
#   - Deletes tasks in folder but missing from list
# All tasks must belong to the same folder, and the folder can't be root
# (too risky to delete unrelated tasks)
scheduler.sync(tasks)
```

## Tips

- You'll likely need administrator privileges to interact with the Task Scheduler;
- A task path uses `\` as separator and starts with `\`. Ex : `\myApp\task`;
- Action paths can use `\` or `/` as separators. When in doubt, use absolute paths;
- To run a task without storing credentials and without requiring the user to be logged-in (on a server, for example), the logging type `Service For User` (`S4U`) worked best for me. This is the default used, but this can be changed according to your needs;
- You'll likely want to template your task definitions. No need to use third-party librairies, just use Python's str.format or String.Template.

## Reference

- [Task Scheduler Object](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-objects)
- [Time Duration String](https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-executiontimelimit#property-value)
