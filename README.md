# Excel Calculator PoC

This project is a **proof of concept (PoC)** that demonstrates how to use Java and Apache POI to evaluate custom Excel formulas such as `XIRR` and `Goal Seek`, directly in-memory without user interaction. It showcases how to embed custom functions into Excel files and calculate them programmatically.

## 🚀 Features

- Evaluate Excel formulas programmatically using Apache POI.
- Support for custom Excel functions:
  - `goalSeek`
  - `myXIRR`
- Generates `.xlsx` files with evaluated formulas.
- Measures execution time for performance profiling.

## 🛠 Technologies Used

- **Java 21+**
- **Apache POI**
- Custom UDFs (User Defined Functions) for Excel.

## 📦 Getting Started

### Clone the repository

```bash
git clone https://github.com/mengzoltan/excel-calculator-poc.git
cd excel-calculator-poc
```

### Build & Run

Use your favorite Java IDE or the command line:

```bash
mvn clean install
```

> ⚠️ Ensure you have the required Apache POI dependencies on the classpath (e.g., `poi`, `poi-ooxml`, `poi-ooxml-schemas`).

### Output

The program generates two Excel files in the root directory:

- `normal_results.xlsx` – evaluated using the full formula set (including Goal Seek)
- `xirr_result.xlsx` – evaluated using only the `XIRR` function

Execution time is printed to the console.

## 🧪 Example Output

```
---------GOAL SEEK -----------
2025-06-21T11:40:04.624686Z main ERROR Log4j API could not find a logging provider.
=========2025-06-21T13:40:05.108198
2025-06-21T13:40:05.141636
A megoldás: 310.687856541074
0.05827879436139105
2025-06-21T13:40:05.316186
0.05827879436139105
=========2025-06-21T13:40:05.326637
218 ms

---------XIRR-----------
=========2025-06-21T13:40:05.436712
0.05827879436139105
=========2025-06-21T13:40:05.447350
10 ms
```

## 📁 Project Structure

```text
src/main/java/studio/intuitech/
├── MainExcelCalculator.java
├── event/
│   ├── Event.java
│   ├── EventDispatcher.java
│   └── Handler.java
├── functions/
│   ├── CalculationExcelTemplateFactory.java
│   ├── CalculationExcelType.java
│   ├── goalseek/
│   │   ├── GoalSeekContext.java
│   │   └── GoalSeekFunction.java
│   ├── newtonsmethod/
│   │   ├── NewtonConfig.java
│   │   └── NewtonsMethod.java
│   └── xirr/
│       ├── Xirr.java
│       ├── XirrContext.java
│       └── XirrFunction.java
```

## 📌 Notes

- Excel file templates are loaded via `CalculationExcelTemplateFactory` as a byte array.
- Custom functions are registered through Apache POI’s UDF (User Defined Function) API.

## ⚖️ License

This project is licensed under the **Apache License 2.0**. See the [LICENSE](LICENSE) file for details.

## ❗ Disclaimer

This proof of concept is intended for **educational and research purposes only**. It should not be used in production systems without proper validation and security review. Use at your own risk.
