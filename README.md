# Bank Statement Analyser

## Getting Started
Please Note, the following steps are only applicable for MacOS and Linux. Any desktop environment excluding those will need additional set-up and configurations.

### Prerequisites
Ensure you have the following installed on your machine:
- Git
- Homebrew (for macOS/Linux)
- Python 3
- pip (Python package installer)
- Tesseract OCR
- Bun (TypeScript Runtime & package manager)

### Installation

1. **Clone the repository**
    ```sh
    git clone <https://github.com/pranavnahar/BSA.git>
    cd <BSA>
    ```

2. **Install Homebrew (for macOS/Linux)**
    ```sh
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    ```

3. **Install Python 3 and pip3 with Homebrew**
    ```sh
    brew install python
    ```

4. **Install Tesseract OCR globally with Homebrew**
    ```sh
    brew install tesseract
    ```

5. **Install project dependencies**
    ```sh
    pip3 install -r requirements.txt
    ```

6. **Install Bun**
    ```sh
    curl -fsSL https://bun.sh/install | sudo bash
    ```

### Configuration

1. **Update `data_folder_path` in `index.py`**
   - Open `index.py` in your preferred text editor.
   - Locate the following lines:
     ```python
     data_folder_path = r"/Users/oeuvars/Documents/ekarth-ventures/BSA/bank-statement/LANDCRAFT-RECREATIONS"
     party_name = "LANDCRAFT RECREATIONS"
     ```
   - Replace the value of `data_folder_path` with the absolute path of the folder `LANDCRAFT-RECREATIONS` on your system. Anything else with result in a error.

2. **Update `ner_model` path in `report_function.py`**
   - Open `report_function.py` in your preferred text editor.
   - Locate line 234:
     ```python
     nlp = spacy.load(r"C:\Users\Lenovo\OneDrive\Desktop\Folders\NaharOm\BSA\Main_Project\ner_model")
     ```
   - Replace the value of the path with the absolute path of the `ner_model` folder on your system. Anything else with result in a error.

### Usage

To start the project, follow these steps:

1. **Install bun dependencies**
    ```bash
    bun install
    ```

2. **Start the Bun web server**
    ```sh
    bun run server.ts
    ```

3. **Invoke the main function**
   - The main function `index.py` is invoked by the Bun web server file `server.ts` when `run-bsa` is hit in the URL.
   - By default, bun runs on port 3000.
   - Access the URL endpoint to start BSA (Bank Statement Analyzer):
     ```
     http://localhost:3000/run-bsa
     ```
Please note, the process takes a lot of time, ~2 minutes or more to generate all the files which can be found under `/bank-statement/LANDCRAFT-RECREATIONS` as xlsx files.


