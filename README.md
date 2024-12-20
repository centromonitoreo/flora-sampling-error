# Application for Sampling Error Calculation for Biotic Environment - Flora

This repository contains a desktop application developed in the Python programming language to evaluate the sampling error of forest harvesting areas in the biotic environment - Flora. This application allows evaluations to be performed from the geodatabase model (GDBs) or by processing information from an Excel file.

## Features

To use the application via the Excel file option, the data must be included in a file with the following columns. The names of the columns must not be modified, as these strings are the values the application will search for to perform the calculations:

## Example of Data Table

| BIOMA  | N_COBERT          | ID_MUEST | VOL_TOTAL  | AREA_UM_ha |
|--------|-------------------|----------|------------|------------|
| 10601  | Arbustal Abierto  | A14      | 0.44060165 | 0.01       |
| 10601  | Arbustal Abierto  | A20      | 0.46352997 | 0.1        |
| 10601  | Arbustal Abierto  | P010     | 0.28712812 | 0.05       |
| 10601  | Arbustal Abierto  | P015     | 0.30072793 | 0.05       |
| 10601  | Arbustal Abierto  | P017     | 0.29180920 | 0.1        |

## Requirements

- Python 3.x  

Install dependencies:

```bash
pip install -r requirements.txt
```

## Installation

### From Python

1. Clone the repository:

    ```bash
    git clone https://github.com/centromonitoreo/flora-sampling-error.git
    ```

2. Run the main script with the graphical interface:

    ```bash
    python src/main.py
    ```

## Contributions

Contributions are welcome. Please follow the repository guidelines for more details.

