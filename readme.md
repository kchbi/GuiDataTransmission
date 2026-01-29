# Match Network Analyzer

This application allows for testing and flashing various matching network assemblies.

## Prerequisites

Before you begin, you must install the following software on your Windows machine:

1.  **Python 3.x:** If you do not have Python, download it from [python.org](https://www.python.org/downloads/). During installation, make sure to check the box that says **"Add Python to PATH"**.

2.  **STM32CubeProgrammer:** This software is required for flashing firmware to the hardware.
    *   Download it directly from the [STMicroelectronics official website](https://www.st.com/en/development-tools/stm32cubeprog.html).
    *   Run the installer and follow the on-screen instructions. Please use the default installation location.

## Setup Instructions

Once the prerequisites are installed, you need to configure your system to recognize the flashing tool. We have provided a script to do this automatically.

1.  **Run the Environment Setup Script:**
    *   Navigate to the application folder.
    *   Double-click the `setup_env.py` file.
    *   A Windows security prompt (UAC) will appear asking for permission. You **must click "Yes"** for the script to work.
    *   A black console window will appear, show its progress, and inform you when it's complete. You can then press Enter to close it.

    *This step only needs to be done once.*

## Running the Application

After completing the setup, you are ready to run the analyzer.

1.  **Install Python Dependencies:**
    *   Open a Command Prompt or PowerShell.
    *   Navigate to the application folder.
    *   Run the command: `pip install -r requirements.txt`

2.  **Launch the Program:**
    *   In the same application folder, double-click `main.py` or run the command: `python main.py`

3  **Changes Required:** 
    *   Serial Number for Assemblies 
    *   Assembly for Part Number
    *   