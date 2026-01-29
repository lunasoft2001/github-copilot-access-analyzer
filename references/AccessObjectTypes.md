# Access Object Types Reference

Reference guide for Microsoft Access object types and their properties.

## Object Type Constants

### Application.SaveAsText / LoadFromText

| Constant | Value | Description |
|----------|-------|-------------|
| acTable | 0 | Table |
| acQuery | 1 | Query |
| acForm | 2 | Form |
| acReport | 3 | Report |
| acMacro | 4 | Macro |
| acModule | 5 | Module |
| acDataAccessPage | 6 | Data Access Page (deprecated) |
| acServerView | 7 | Server View |
| acDiagram | 8 | Database Diagram |
| acStoredProcedure | 9 | Stored Procedure |
| acFunction | 10 | Function |

### VBA Component Types

| Type | Value | Description |
|------|-------|-------------|
| vbext_ct_StdModule | 1 | Standard Module (.bas) |
| vbext_ct_ClassModule | 2 | Class Module (.cls) |
| vbext_ct_MSForm | 3 | UserForm (not used in Access) |
| vbext_ct_ActiveXDesigner | 11 | ActiveX Designer |
| vbext_ct_Document | 100 | Document Module (Forms/Reports in Access) |

### DAO Field Types

| Constant | Value | Description | Size |
|----------|-------|-------------|------|
| dbBoolean | 1 | Yes/No | 1 bit |
| dbByte | 2 | Byte | 1 byte |
| dbInteger | 3 | Integer | 2 bytes |
| dbLong | 4 | Long Integer | 4 bytes |
| dbCurrency | 5 | Currency | 8 bytes |
| dbSingle | 6 | Single | 4 bytes |
| dbDouble | 7 | Double | 8 bytes |
| dbDate | 8 | Date/Time | 8 bytes |
| dbBinary | 9 | Binary | Variable |
| dbText | 10 | Text | Variable (up to 255) |
| dbLongBinary | 11 | OLE Object | Variable |
| dbMemo | 12 | Memo/Long Text | Variable |
| dbGUID | 15 | Replication ID | 16 bytes |
| dbBigInt | 16 | Big Integer | 8 bytes |
| dbVarBinary | 17 | VarBinary | Variable |
| dbChar | 18 | Char | Variable |
| dbNumeric | 19 | Numeric | Variable |
| dbDecimal | 20 | Decimal | Variable |
| dbFloat | 21 | Float | 4 bytes |
| dbTime | 22 | Time | 8 bytes |
| dbTimeStamp | 23 | Time Stamp | 8 bytes |

## Table Attributes

| Attribute | Value | Description |
|-----------|-------|-------------|
| dbAttachExclusive | 0x10000 | Opened exclusively |
| dbAttachSavePWD | 0x20000 | Save password with connection |
| dbSystemObject | 0x80000002 | System object |
| dbAttachedTable | 0x40000000 | Attached (linked) table |
| dbAttachedODBC | 0x20000000 | ODBC linked table |
| dbHiddenObject | 1 | Hidden object |

## Common Collections

### CurrentProject Collections
- `AllForms` - All forms in database
- `AllReports` - All reports in database
- `AllMacros` - All macros in database
- `AllModules` - All standard modules in database
- `AllDataAccessPages` - All data access pages (deprecated)

### DAO Database Collections
- `TableDefs` - All table definitions
- `QueryDefs` - All query definitions
- `Relations` - All relationships
- `Containers` - System containers

### VBProject Collections
- `VBComponents` - All VBA components (modules, classes, forms, reports)
- `References` - All VBA references

## Export/Import Patterns

### Identifying User Objects

**User Tables:**
```vba
Function IsUserTable(tbl As DAO.TableDef) As Boolean
    ' Exclude system and hidden objects
    If (tbl.Attributes And (dbSystemObject Or dbHiddenObject)) <> 0 Then
        IsUserTable = False
        Exit Function
    End If
    
    ' Exclude MSys* and USys* tables
    Dim name As String
    name = UCase$(tbl.Name)
    IsUserTable = Not (Left$(name, 4) = "MSYS" Or Left$(name, 4) = "USYS")
End Function
```

**User Queries:**
```vba
Function IsUserQuery(qry As DAO.QueryDef) As Boolean
    ' Exclude temporary queries (~sq_*) and system queries
    Dim name As String
    name = qry.Name
    IsUserQuery = Not (Left$(name, 4) = "~sq_" Or Left$(UCase$(name), 4) = "MSYS")
End Function
```

### VBA Component Names

Access VBA components follow naming patterns:
- Forms: `Form_<FormName>`
- Reports: `Report_<ReportName>`
- Standard Modules: `<ModuleName>`
- Class Modules: `<ClassName>`

## COM Automation

### Creating Access Application

```powershell
$access = New-Object -ComObject Access.Application
$access.Visible = $false
```

### Opening Database

```powershell
# Non-exclusive (for reading/exporting)
$access.OpenCurrentDatabase($path, $false)

# Exclusive (for importing/modifications)
$access.OpenCurrentDatabase($path, $true)
```

### Accessing VBA Project

```powershell
$vbProject = $access.VBE.ActiveVBProject
```

### Running VBA Procedures

```powershell
# Without parameters
$access.Run("ProcedureName")

# With parameters
$access.Run("ProcedureName", $param1, $param2)
```

### Importing VBA Modules

```powershell
$vbProject.VBComponents.Import($filePath)
```

### Removing VBA Modules

```powershell
$component = $vbProject.VBComponents.Item("ModuleName")
$vbProject.VBComponents.Remove($component)
```

## File Extensions

| Extension | Description |
|-----------|-------------|
| .accdb | Access 2007+ database |
| .mdb | Access 2003 database |
| .accde | Access 2007+ compiled database (VBA compiled, no editing) |
| .mde | Access 2003 compiled database |
| .accdr | Access 2007+ runtime database |
| .laccdb | Lock file (multi-user) |
| .ldb | Lock file (Access 2003) |

## Common Errors

| Error | Code | Description |
|-------|------|-------------|
| Database in use | 3045 | Database is locked by another user |
| Permission denied | 3051 | Cannot open file exclusively |
| VBA access denied | -2147467259 | "Trust access to VBA project object model" not enabled |
| Object not found | 3265 | Item not found in collection |
| Type mismatch | 13 | Data type mismatch |
