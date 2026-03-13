# Java-Excel-API

A Java Swing desktop application for importing, viewing, and exporting Excel (`.xlsx`) files, with MySQL-backed user authentication. The application supports three independent, company-branded modules: **Ooredoo**, **Orange**, and **Telecom Algeria**.

---

## Features

- **Welcome screen** — select which company module to launch
- **User authentication** — login and registration with MD5 password hashing stored in a MySQL database
- **Excel import** — load `.xlsx` files via Apache POI; supports merged/grouped column headers
- **Data display** — renders Excel table data in an interactive Swing table (JBroTable) that preserves grouped header structure
- **Excel export** — export the currently displayed data back to a formatted `.xlsx` file with styled headers
- **Per-company branding** — each module (Ooredoo, Orange, Telecom) has its own login, registration, and main frame

---

## Project Structure

```
Java-Excel-API/
├── src/
│   ├── Welcome/
│   │   └── Welcome.java              # Application entry point (company selector)
│   └── app/
│       ├── ooreedo/                  # Ooredoo module
│       │   ├── login.java            # Login UI
│       │   ├── register.java         # Registration UI
│       │   ├── testframe.java        # Main application frame (Import/Export menu)
│       │   ├── Import.java           # Excel import logic (Apache POI)
│       │   ├── export.java           # Excel export logic (Apache POI)
│       │   ├── Display.java          # Table rendering with grouped headers
│       │   ├── ColumnGroup.java      # Column group model
│       │   ├── GroupableTableHeader.java    # Custom table header
│       │   └── GroupableTableHeaderUI.java  # Custom table header UI
│       ├── orange/                   # Orange module (same structure)
│       └── telecom/                  # Telecom Algeria module (same structure)
├── lib/                              # All dependency JARs
├── img/                              # Company logos (ooredoo.png, orange.png, telecom.png)
├── bin/                              # Compiled .class files
└── README.md
```

---

## Technologies & Dependencies

| Library | Version | Purpose |
|---|---|---|
| Apache POI | 4.1.0 | Read and write Excel `.xlsx` files |
| MySQL Connector/J | 8.0.29 | Database connectivity for authentication |
| JBroTable | 1.1.3 | Swing table component with grouped/merged column headers |
| Commons Codec | 1.12 | Encoding utilities |
| Commons Collections | 4.3 | Extended Java collections |
| Commons Compress | 1.18 | Compression support (used by POI) |
| Commons Logging | 1.2 | Logging facade |
| Commons Math | 3.6.1 | Math utilities |
| JAXB API / Core / Impl | 2.3.x | XML binding (required by POI on Java 9+) |
| XMLBeans | 3.1.0 | XML processing (used by POI OOXML) |
| CurvesAPI | 1.06 | Chart curve support (POI dependency) |
| log4j | 1.2.17 | Logging |
| JUnit | 4.12 | Unit testing |

All JARs are included in the `lib/` directory.

---

## Prerequisites

- **Java JDK 8** or higher
- **MySQL** database server running locally
- An IDE such as **Eclipse** or **IntelliJ IDEA** (recommended for classpath setup), or the JDK CLI tools

---

## Database Setup

Each module uses a MySQL database for user authentication. Before running the application, create the required database and users table. Example SQL:

```sql
CREATE DATABASE telecom_db;
USE telecom_db;

CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL UNIQUE,
    password VARCHAR(32) NOT NULL,   -- MD5 hash
    email VARCHAR(150) NOT NULL
);
```

> Repeat for each module's database (e.g., `ooredoo_db`, `orange_db`). Update the JDBC connection strings inside each module's `login*.java` and `register*.java` files to match your database name, host, port, and credentials.

---

## Build & Run

### Using Eclipse

1. Clone or download the repository.
2. Open Eclipse and select **File → Import → Existing Projects into Workspace**, then point to the project folder.
3. Right-click the project → **Build Path → Configure Build Path → Libraries → Add JARs**, and add all JARs from the `lib/` directory.
4. Run `Welcome.java` as a Java application.

### Using the Command Line

```bash
# From the project root
javac -cp "lib/*" -d bin src/Welcome/Welcome.java src/app/ooreedo/*.java src/app/orange/*.java src/app/telecom/*.java

java -cp "bin;lib/*" Welcome.Welcome
```

> On Linux/macOS, replace `;` with `:` in the classpath.

---

## Usage

1. **Launch** the application — the Welcome screen displays logos for the three companies.
2. **Click** a company logo to open that module's login screen.
3. **Log in** with existing credentials, or click **Register** to create a new account.
4. After login, the **main frame** opens with a menu bar:
   - **File → Import → Template with data** — opens a file chooser to select an `.xlsx` file; the table is populated with the Excel data, preserving merged header groups.
   - **File → Export** — saves the current table data to a new `.xlsx` file in a location you choose.

---

## Architecture Overview

```
Welcome (entry point)
    │
    ├── login / register  ──► MySQL (MD5 hashed passwords)
    │
    └── testframe (main frame)
            │
            ├── Import ──► Apache POI ──► reads .xlsx
            │       │
            │       └── Display ──► JBroTable (grouped headers)
            │
            └── export ──► Apache POI ──► writes .xlsx
```

---

## License

All rights reserved by the original authors.
