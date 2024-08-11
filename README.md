# Lecture Map Plotter
A graphical user interface (GUI) application designed to visualize and manage lecture and event data on an interactive map. This tool allows users to input, edit, and plot the locations of lectures or events, offering an intuitive approach to tracking educational activities geographically.

---
## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

---
## Features
- Address Geocoding: Lookup addresses and retrieve their geographic coordinates.
- Interactive Map Plotting: Visualize event locations on a map with customizable views.
- Data Management: Add, edit, and remove event data stored in Excel files.
- Custom Views: Create and manage personalized map views for different regions.
- Statistics Calculation: Compute and display event statistics based on the provided data.

---
## Requirements
- Pyenv: (Recommended) A tool for managing Python versions installed on your machine.
- Python 3.11.x: Ensure that you have Python version 3.11.x installed on your system. You can download it from the official Python website, or if using Pyenv/Pyenv-win, you can install it via the Pyenv interface.
- Poetry: A dependency management tool for Python. 

---
### Pyenv-win
If you're using Windows and would like to use Pyenv to manage your Python installs:

#### Pyenv-win install on Windows (Powershell)

```bash
Invoke-WebRequest -UseBasicParsing -Uri "https://raw.githubusercontent.com/pyenv-win/pyenv-win/master/pyenv-win/install-pyenv-win.ps1" -OutFile "./install-pyenv-win.ps1"; &"./install-pyenv-win.ps1"
```
If you are getting any UnauthorizedAccess error as below then start Windows PowerShell with the “Run as administrator” option and run the code below and re-run the Invoke-WebRequest installation command.
```bash
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```
[Pyenv-win Documentation](https://pyenv-win.github.io/pyenv-win/docs/installation.html)
---

---
### Poetry
If you don't have Poetry installed, you can install it using the following command:

#### Poetry install on Linux, macOS, Windows(WSL)
```bash
curl -sSL https://install.python-poetry.org | python3 -
```
#### Poetry install on Windows (Powershell)
```bash
(Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | py -
```
For more detailed instructions and alternative installation methods, refer to the Poetry documentation.

[Poetry Documentation](https://python-poetry.org/docs/)
---

---
## Installation

Follow these steps to set up the project on your local machine:

### 1. Clone the repository

Begin by cloning the repository to your local machine:

```bash
git clone https://github.com/Shinekaze/ClimatePlotter.git
cd ClimatePlotter
```

### 2. Install Dependencies

Navigate to the project's directory and use Poetry to install all the required dependencies:

```bash
poetry install
```
This command will create a virtual environment and install all necessary packages as specified in the pyproject.toml file.

### 3. Activate the Virtual Environment

Before running the application, activate the virtual environment created by Poetry:

```bash
poetry shell
```

### 4. Run the Application

With the virtual environment activated, you can now start the application:

```bash
python ClimatePlotter3.py
```

---
## Usage
Once the application is running, you can perform the following actions:

#### Adding Event Data

- Input Details: Enter event specifics such as date, name, address, city, state, postal code (PLZ), number of tables, and participants.
- Lookup Address: Utilize the 'Lookup Address' feature to automatically fetch geographic coordinates based on the provided address.
- Save Data: After verifying the details, save the event data to the Excel file.

#### Managing Views

- Create Custom Views: Define personalized map views by specifying latitude, longitude, and bounding box coordinates.
Save and Manage Views: Store these views for easy access in future sessions. You can also edit or delete existing views as needed.

#### Plotting Events

- Select a View: Choose a desired view from the list to plot events on the map.
Preview Map: Use the 'Pre-View' feature to see how the map will appear before finalizing.
Recalculating Statistics

- After adding or modifying events, employ the 'Recalculate Statistics' feature to update and display the latest event statistics.

---
## License
This project is licensed under the MIT License. You are free to use, modify, and distribute this software in accordance with the license terms.

---
## Acknowledgments
Pyside6: For the GUI framework.
Pandas: For data manipulation and analysis.
OpenStreetMap: For geocoding services and map data.
Poetry: For dependency management and packaging.

