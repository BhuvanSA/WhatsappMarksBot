# Marks Sender Bot

Marks Sender Bot is a Python-based GUI application that automates the process of sending student marks to parents via WhatsApp. It reads data from an Excel file and uses Selenium to interact with WhatsApp Web for message delivery.

![Marks Sender Bot GUI](/data/Preview.png)

## Features

-   User-friendly GUI built with Tkinter and ttkbootstrap
-   Excel file parsing for student data
-   Automated WhatsApp messaging using Selenium
-   Progress tracking and status updates
-   Customizable message templates

## Installation

1. Clone the repository:

    ```
    git clone https://github.com/BhuvanSA/WhatsappMarksBot.git
    cd marks-sender-bot
    ```

2. Create a virtual environment (optional but recommended):

    ```
    conda create -n whatsappbot python=3.11
    conda activate whatsappbot
    ```

3. Install the required dependencies:

    ```
    pip install -r requirements.txt
    ```

4. Ensure you have Chrome installed, as the bot uses ChromeDriver for Selenium.

## Usage

1. Run the main application:

    ```
    python src/main.py
    ```

2. Use the GUI to:

    - Select the Excel file containing student data
    - Choose the appropriate sheet and internal assessment
    - Set the range of students to process
    - Enter the mentor's name
    - Start the sending process

3. The application will open WhatsApp Web. Scan the QR code to log in.

4. The bot will automatically send messages to the specified range of students.

5. Monitor the progress and status in the GUI table.

<!-- ![Marks Sender Bot in Action](path/to/bot_in_action.gif) -->

## Project Structure

-   `src/`
    -   `main.py`: Entry point of the application
    -   `gradebook/gradebook.py`: Main GUI and application logic
    -   `excelManager.py`: Handles Excel file operations
    -   `seleniumManager.py`: Manages Selenium WebDriver for WhatsApp interaction
    -   `messageGenerator.py`: Generates message content

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgements

-   [ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap) for the modern GUI elements
-   [Selenium](https://www.selenium.dev/) for web automation
-   [openpyxl](https://openpyxl.readthedocs.io/) for Excel file handling

## Disclaimer

This project is for educational purposes only. Use it responsibly and in compliance with WhatsApp's terms of service.
