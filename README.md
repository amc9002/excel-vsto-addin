# Excel JSON Exporter (VSTO Add-in)

Excel VSTO add-in for reading structured data from Excel worksheets and exporting it to JSON (and other formats).

The project demonstrates clean separation between Excel/VSTO integration, business logic, and testable core services.

---

## ✨ Features

- Excel VSTO add-in (.NET Framework 4.8)
- Reads data from Excel worksheets (UsedRange)
- Maps rows to strongly typed domain models
- Exports data to:
  - JSON (via Newtonsoft.Json)
- Unit-tested core logic (xUnit)
- Clean architecture: VSTO layer → Core → Tests

---

## 🏗 Project Structure

ExcelAddIn1
├── ExcelAddIn1VSTO # Excel VSTO add-in (UI, Ribbon, COM interaction)
├── ExcelAddIn1.Core # Core business logic (no Excel dependencies)
│ ├── Models # Domain models (Person, etc.)
│ └── Services # Mapping & export services
└── ExcelAddIn1.Tests # Unit tests (xUnit)


---

## 🧠 Architecture Overview

**VSTO Layer**
- Handles Excel COM interaction
- Reads worksheet data
- Delegates logic to Core services

**Core Layer**
- No dependency on Excel or COM
- Contains:
  - Row → model mapping logic
  - Export services (JSON, extensible)
- Fully unit-testable

**Tests**
- xUnit-based tests
- Cover:
  - Row-to-model mapping
  - Edge cases (nulls, empty values, invalid data)
  - JSON serialization and file output

---

## 🧪 Example: Mapping Excel Row to Model

```csharp
object[] row = { "Alice", "30", "Minsk" };

var mapper = new PersonMapper();
Person person = mapper.MapRowToPerson(row);

📤 JSON Export Example

var exporter = new JsonExportService();
exporter.Export(people, "output.json");

Produces formatted JSON output:

[
  {
    "name": "Alice",
    "age": 30,
    "city": "Minsk"
  }
]

✅ Testing

xUnit

Tests focus on logic, not Excel COM

Excel-specific code is isolated from tests

[Fact]
public void MapRow_Should_Map_All_Fields()
{
    object[] row = { "Alice", "30", "Minsk" };

    var person = new PersonMapper().MapRowToPerson(row);

    Assert.Equal("Alice", person.Name);
    Assert.Equal(30, person.Age);
    Assert.Equal("Minsk", person.City);
}

🛠 Technologies

C#

.NET Framework 4.8

Excel VSTO

COM Interop

Newtonsoft.Json

xUnit

🚀 Notes

The architecture allows easy extension:

CSV export

Additional domain models

Alternative data sources

Core logic is reusable outside Excel if needed

📄 License

This project is provided for demonstration and educational purposes.