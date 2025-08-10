# Anki Note Creator

This is a simple desktop application for Windows that allows you to quickly create new notes in your Anki decks directly from your clipboard content, including rich text formatting and images.

---

## Demo



https://github.com/user-attachments/assets/fccf3520-7520-4a2b-9b45-ad30a186f264



## Features

* **Copy content from anywhere:** You want to revise the note, article or content from anywhere? Great, this application is for you, even though it does not create questions automatically, which you don't really need when you want complete content to be bookmarked, utilize Anki's active recalling technique, without any chunks of questions.
* **Automatic Formatting:** Converts clipboard text with formatting (bold, italic, etc.) and images into Anki-ready HTML.
* **Image Handling:** Automatically scales and saves images from the clipboard to your Anki media library.
* **Deck Management:** Refresh, add, edit, and delete decks directly within the application.
* **Standalone Executable:** Provides a pre-built `.exe` file for users who don't want to deal with Python.

## How to Use the Pre-built Executable

1.  Go to the **[Releases page](https://github.com/YOUR_USERNAME/Anki-Note-Creator/releases)** of this repository.
2.  Download the latest `Anki_Note_Creator.zip` file.
3.  Extract the contents to a folder on your computer.
4.  Double-click `Anki_Note_Creator.exe` to run the application.

> **Note:** You must have the Anki desktop application and the AnkiConnect add-on running for this tool to work.

## For Developers

### Prerequisites

* Python 3.x
* [Anki Desktop Application](https://apps.ankiweb.net/)
* [AnkiConnect Add-on](https://ankiweb.net/shared/info/2055492159)

### Installation

1.  Clone the repository:
    ```bash
    git clone [https://github.com/YOUR_USERNAME/Anki-Note-Creator.git](https://github.com/YOUR_USERNAME/Anki-Note-Creator.git)
    cd Anki-Note-Creator
    ```
2.  Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

### Building the Executable

You can create your own executable using PyInstaller.
```bash
pip install pyinstaller
pyinstaller --onefile --windowed "Transfer Any Article or Note to Anki.py"
