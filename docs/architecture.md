# System Architecture

## Overview

The system is a desktop CRM application designed to manage tasks and interactions in an operational environment. System works offline - no internet needed, no subscription or any other payments, complete free.

---

## Tech Stack

- Python — core application logic  
- SQLite — database  
- customtkinter / tkinter — user interface  

---

## Architecture Layers

### 1. UI Layer

Responsible for:
- displaying task lists
- forms for creating/editing tasks
- filters and dashboards

Key components:
- task list (TreeView)
- task form
- filters and search
- dashboard panel

---

### 2. Application Layer (Logic)

Responsible for:
- task lifecycle management
- status calculation
- filtering and prioritization
- business rules

Examples:
- updating task status based on dates
- activating/deactivating interactions
- calculating overdue tasks

---

### 3. Data Layer

- SQLite database
- structured tables:
  - Persons
  - Interactions
  - Tasks
  - Dictionaries (responsibles, circles)

---

## Key Design Decisions

- lightweight architecture for fast deployment  
- local database (no infrastructure dependency)  
- modular structure for future scalability  
- focus on usability over complexity  

---

## Why This Architecture

The system was designed to:

- be quickly implemented in a real environment  
- require minimal setup  
- be understandable for non-technical users  
- support process transformation without heavy IT involvement  

---

## Limitations (honest)

- desktop-only (no web access)  
- single-user focus  
- limited integration capabilities  

---

## Future Improvements

- web version  
- multi-user support  
- API integrations  
- role-based access  
