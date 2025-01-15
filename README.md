# Accounting Program

A Python-based accounting tool designed for ease of use and functionality. Developed by Gabriele Lo Cascio, this software replaces legacy MS-DOS accounting programs with a modern, feature-rich interface.

## Features

- **Main Class (`Cashier`)**: Acts as the core of the application, managing other components.
- **Menu and Options**:
  - Import data from Excel files.
  - Edit client lists and print options.
  - Create backups.
- **Input and Data Management**:
  - Daily input of transactions with support for manual and Excel-based entries.
  - Automatic balance updates.
- **Reporting**:
  - Generate summaries of storage, daily entries, and individual accounts.
  - Export data in various formats.

## Classes Overview

- **`Cashier`**: Main container managing the flow of the application.
- **`MenuP`**: Main navigation hub.
- **`Input`**: Interface for entering daily transactions.
- **`PrintStorage`, `PrintDay`, `PrintSingle`**: Different views for summarizing and analyzing data.
- **`EditValue`**: Modify and remove entries.
- **`Reset`**: Clear zero-balance accounts.
- **`UndoLastImport`**: Remove the most recently added data.
- **`Transmission`**: Migrate data between accounts.

## Getting Started

### Prerequisites

- Python 3.x
- Required Libraries: `tkinter`, `ttkthemes`, `tkcalendar`, `pandas`, `numpy`, `matplotlib`, `seaborn`

### Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/Gabro29/Accounting
    ```
2. Run the program:
    ```bash
    python Accounting.py
    ```

## Usage

1. Navigate through the main menu to access different functionalities.
2. Use the Input screen to add daily transactions.
3. Generate reports and backups as needed.

## License

This project is licensed under the [Apache License 2.0](http://www.apache.org/licenses/LICENSE-2.0). 

If you use or distribute this code (with or without modifications), you are required to:

1. Proper attribution to the original author. Retain the following attribution in documentation or credits:
   ```
	This project is based on code originally developed by Gabriele Lo Cascio.
	The original repository is available at https://github.com/Gabro29/Accounting
   ```
2. Include the original `LICENSE` files in your distribution.
3. Clearly indicate any modifications made to the original project.
4. Adhere to the terms specified in the [Apache License 2.0](http://www.apache.org/licenses/LICENSE-2.0).

## Contact

For any inquiries or support, reach out to Gabriele Lo Cascio:

- **LinkedIn**: [Gabriele Lo Cascio](https://www.linkedin.com/in/gabriele-locascio)
- **GitHub**: [Gabro29](https://github.com/Gabro29)
- **Fiverr**: [Gabro29](https://it.fiverr.com/gabro_29?up_rollout=true)
- **YouTube**: [Channel](https://www.youtube.com/channel/UCkGvbGqYzDi3lfgtbQ_pngg)
- **Instagram**: [@ga8ro](https://www.instagram.com/ga8ro)
