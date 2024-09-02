"""
Provides a Python interface for working with the Windows Task Scheduler using the `win32com.client` module from `pywin32`.
It allows for the creation, registration, deletion, and synchronization of scheduled tasks on a Windows system.

Key Features:
  - Task Creation: Build task definitions with default and custom attributes from a Python dict.
  - Task Registration: Register tasks with specific logon types, user credentials, and creation flags.
  - Task Deletion: Delete specific tasks or entire folders from the Task Scheduler.
  - Task Synchronization: Sync a list of tasks with the Task Scheduler, ensuring tasks are created, updated, or deleted as needed.

This module makes use of constants and bit flag mappings derived from Microsoft's documentation to interact with various task properties like
triggers, actions, logon types, and run levels. It should be equivalent to Microsoft's XML specs. The Python dict used for task creation should
mirror the task's data model.

This code must be run with administrator privileges to work.


Example Usage:

# Initialize
scheduler = TaskScheduler()

# Load your task data
tasks_definitions = ...  # See README for examples

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
scheduler.sync(tasks)


Additionnal reference:
  - Win32 Task Data Model: https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-objects
  - Time Duration : https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-executiontimelimit#property-value


License:

MIT License

Copyright (c) 2024 Jean-Michel Grenier

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import contextlib
import datetime
import logging
import os
from dataclasses import dataclass
from functools import reduce
from pathlib import Path
from typing import Any

import pywintypes
import win32com.client

logger = logging.getLogger(__name__)

# Mappings for the constants

# https://learn.microsoft.com/en-us/windows/win32/taskschd/trigger-type
TRIGGER_TYPES: dict[str, int] = {
    'EVENT': 0,
    'TIME': 1,
    'DAILY': 2,
    'WEEKLY': 3,
    'MONTHLY': 4,
    'MONTHLYDOW': 5,
    'IDLE': 6,
    'REGISTRATION': 7,
    'BOOT': 8,
    'LOGON': 9,
    'SESSION_STATE_CHANGE': 11,
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/action-type
ACTION_TYPES: dict[str, int] = {
    'EXEC': 0,
    'COM_HANDLER': 5,
    'SEND_EMAIL': 6,
    'SHOW_MESSAGE': 7,
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/taskfolder-registertask
# https://learn.microsoft.com/en-us/windows/win32/taskschd/principal-logontype
LOGON_TYPES: dict[str, int] = {
    'NONE': 0,
    'PASSWORD': 1,
    'S4U': 2,
    'INTERACTIVE_TOKEN': 3,
    'GROUP': 4,
    'SERVICE_ACCOUNT': 5,
    'INTERACTIVE_TOKEN_OR_PASSWORD': 6,
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/principal-runlevel
TASK_RUNLEVEL_TYPE: dict[str, int] = {'LUA': 0, 'HIGHEST': 1}

# Bit Flags Mappings

# https://learn.microsoft.com/en-us/windows/win32/taskschd/taskfolder-registertask
TASK_CREATION_FLAGS: dict[str, int] = {
    'VALIDATE_ONLY': 0x1,
    'CREATE': 0x2,
    'UPDATE': 0x4,
    'CREATE_OR_UPDATE': 0x6,
    'DISABLE': 0x8,
    'DONT_ADD_PRINCIPAL_ACE': 0x10,
    'IGNORE_REGISTRATION_TRIGGERS': 0x20,
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/monthlydowtrigger-daysofweek
DAYSOFWEEK: dict[str, int] = {
    'Sunday': 0x1,
    'Monday': 0x2,
    'Tuesday': 0x4,
    'Wednesday': 0x8,
    'Thursday': 0x10,
    'Friday': 0x20,
    'Saturday': 0x40,
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/monthlytrigger-daysofmonth
DAYSOFMONTH: dict[str, int] = {
    '1': 0x1,
    '2': 0x2,
    '3': 0x4,
    '4': 0x8,
    '5': 0x10,
    '6': 0x20,
    '7': 0x40,
    '8': 0x80,
    '9': 0x100,
    '10': 0x200,
    '11': 0x400,
    '12': 0x800,
    '13': 0x1000,
    '14': 0x2000,
    '15': 0x4000,
    '16': 0x8000,
    '17': 0x10000,
    '18': 0x20000,
    '19': 0x40000,
    '20': 0x80000,
    '21': 0x100000,
    '22': 0x200000,
    '23': 0x400000,
    '24': 0x800000,
    '25': 0x1000000,
    '26': 0x2000000,
    '27': 0x4000000,
    '28': 0x8000000,
    '29': 0x10000000,
    '30': 0x20000000,
    '31': 0x40000000,
    'Last': 0x80000000,  # Seems buggy with pywin32; 0x80000000 is max int32
}

# https://learn.microsoft.com/en-us/windows/win32/taskschd/monthlytrigger-monthsofyear
MONTHSOFYEAR: dict[str, int] = {
    'January': 0x1,
    'February': 0x2,
    'March': 0x4,
    'April': 0x8,
    'May': 0x10,
    'June': 0x20,
    'July': 0x40,
    'August': 0x80,
    'September': 0x100,
    'October': 0x200,
    'November': 0x400,
    'December': 0x800,
}

MAPPINGS: dict[str, dict[str, int]] = {'LogonType': LOGON_TYPES, 'RunLevel': TASK_RUNLEVEL_TYPE}

BITFLAG_MAPPINGS: dict[str, dict[str, int]] = {
    'DaysOfWeek': DAYSOFWEEK,
    'DaysOfMonth': DAYSOFMONTH,
    'MonthsOfYear': MONTHSOFYEAR,
}


def default_author() -> str:
    """Identical to the default used by the Task Scheduler's wizard"""
    domain = os.environ['USERDOMAIN']
    author = os.getlogin()
    return f'{domain}\\{author}' if domain else author


TASK_DEFAULTS: dict[str, Any] = {
    'RegistrationInfo': {'Author': default_author(), 'Date': lambda: datetime.datetime.now().isoformat()},
    'Settings': {'ExecutionTimeLimit': 'P1D'},
    'Principal': {'ID': 'Author', 'UserId': default_author(), 'LogonType': 'S4U', 'RunLevel': 'HIGHEST'},
}


def filter_keys(d: dict, keys: list) -> dict:
    return {k: v for k, v in d.items() if k not in keys}


def flag_value(li: list[str | int], mapping: dict[str, int]) -> int:
    return reduce(lambda a, x: a | mapping.get(x, x), li, 0)  # type: ignore


def task_path(path: str) -> Path:
    path = path.replace('/', '\\')
    return Path(path if path.startswith('\\') else f'\\{path}')


def parse_value(key: str, value: Any) -> Any:
    if key in BITFLAG_MAPPINGS and isinstance(value, list):
        # If the value must be computed from a list of flags
        return flag_value(value, BITFLAG_MAPPINGS[key])
    elif key in MAPPINGS and isinstance(value, str):
        # If the value can be str-mapped
        return MAPPINGS[key][value]
    elif isinstance(value, datetime.date):
        # Dates must be converted to iso str
        return value.isoformat()
    elif callable(value):
        # If the value is dynamically generated, ex: creation date
        return value()

    return value


def set_task_attributes(task: win32com.client.CDispatch, attributes: Any) -> None:
    for key, value in attributes.items():
        if isinstance(value, dict):
            # Recursively set attributes for nested dictionaries
            set_task_attributes(getattr(task, key), value)
        elif isinstance(value, list) and key in ('Triggers', 'Actions'):
            # Handle lists for Triggers and Actions collections
            collection = getattr(task, key)
            for item in value:
                if key == 'Triggers':
                    collection_item = collection.Create(TRIGGER_TYPES.get(item['Type'], item['Type']))
                elif key == 'Actions':
                    collection_item = collection.Create(ACTION_TYPES.get(item['Type'], item['Type']))
                set_task_attributes(collection_item, filter_keys(item, ['Type']))
        else:
            # Set value
            setattr(task, key, parse_value(key, value))


@dataclass
class Task:
    path: Path
    definition: win32com.client.CDispatch


class TaskScheduler:
    def __init__(self, task_defaults: dict[str, Any] = TASK_DEFAULTS) -> None:
        self.client: win32com.client.dynamic.CDispatch = win32com.client.Dispatch('Schedule.Service')
        logger.debug('Connecting to Schedule.Service win32com')
        self.client.Connect()
        self.root: win32com.client.CDispatch = self.client.GetFolder('\\')
        self.task_defaults = task_defaults

    def build(self, task_attributes: dict[str, Any]) -> Task:
        """Applies the defaults first, then task_attributes"""
        logger.debug('Building task %s', task_attributes['Path'])
        task_def = self.client.NewTask(0)
        set_task_attributes(task_def, self.task_defaults)
        set_task_attributes(task_def, filter_keys(task_attributes, ['Path']))
        return Task(task_path(task_attributes['Path']), task_def)

    def register(
        self,
        task: Task,
        logonType: str = 'S4U',
        userId: str | None = None,
        password: str | None = None,
        creation_flag: str = 'CREATE_OR_UPDATE',
    ) -> None:
        logger.debug('Registering task %s with flag %s', task.path, creation_flag)
        self._get_and_create_folder(task.path.parent).RegisterTaskDefinition(
            task.path.name,
            task.definition,
            flag_value([creation_flag], TASK_CREATION_FLAGS),
            userId,
            password,
            LOGON_TYPES.get(logonType, logonType),
        )

    def task_exists(self, path: str | Path) -> bool:
        try:
            self.root.getTask(str(path))
            return True
        except pywintypes.com_error:
            return False

    def folder_exists(self, path: str | Path) -> bool:
        try:
            self.root.getFolder(str(path))
            return True
        except pywintypes.com_error:
            return False

    def get_tasks(self, path: str | Path) -> list[win32com.client.CDispatch]:
        return self.root.GetFolder(str(path)).GetTasks(0)

    def delete_task(self, path: str | Path) -> None:
        logger.debug('Deleting task %s', path)
        self.root.DeleteTask(str(path), 0)

    def delete_folder(self, path: str | Path) -> None:
        logger.debug('Deleting folder %s', path)
        self.root.DeleteFolder(str(path), 0)

    def sync(
        self, tasks: list[Task], logonType: str = 'S4U', userId: str | None = None, password: str | None = None
    ) -> None:
        """
        - Creates tasks in input but missing from folder
        - Updates tasks in input already in folder
        - Deletes tasks in folder but missing from input

        The tasks must belong to the same folder and can't be in the root folder
        (since most applications put their tasks there, it woulkd be too dangerous to delete them)
        """
        # Validation
        if len({task.path.parent for task in tasks}) > 1:
            raise AttributeError('The tasks must belong to the same folder')

        folder = tasks[0].path.parent

        if folder == Path('\\'):
            raise AttributeError("The tasks' folder can't be root (there is a high risk of deleting unrelated tasks)")

        # Creation / Update
        for task in tasks:
            self.register(
                task,
                logonType=logonType,
                userId=userId,
                password=password,
                creation_flag='UPDATE' if self.task_exists(task.path) else 'CREATE',
            )

        # Suppression
        folder_task_paths = set(Path(task.Path) for task in self.get_tasks(folder))
        wanted_task_paths = {task.path for task in tasks}
        for path in folder_task_paths - wanted_task_paths:
            self.delete_task(path)

    def _get_and_create_folder(self, path: str | Path) -> win32com.client.CDispatch:
        with contextlib.suppress(pywintypes.com_error):
            self.root.CreateFolder(str(path))
        return self.client.GetFolder(str(path))
