# VBA Refactoring Patterns and Guidelines

Common patterns and best practices for refactoring Access VBA code.

## Code Smells to Look For

### 1. God Procedures
Procedures with 100+ lines that do too many things.

**Fix:** Extract logical blocks into separate functions.

```vba
' Before
Sub ProcessInvoice()
    ' 200 lines of mixed concerns
End Sub

' After
Sub ProcessInvoice()
    ValidateInvoiceData
    CalculateInvoiceTotals
    ApplyDiscounts
    SaveInvoiceToDatabase
    SendInvoiceEmail
End Sub
```

### 2. Magic Numbers
Hard-coded numeric values without explanation.

**Fix:** Use named constants.

```vba
' Before
If status = 3 Then
    ...
End If

' After
Const STATUS_APPROVED As Integer = 3
If status = STATUS_APPROVED Then
    ...
End If
```

### 3. Repeated Code
Same logic duplicated across multiple procedures.

**Fix:** Extract to shared function.

```vba
' Before
Sub Proc1()
    On Error GoTo ErrH
    ' ... logic ...
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
End Sub

Sub Proc2()
    On Error GoTo ErrH
    ' ... logic ...
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
End Sub

' After - Shared error handler
Sub Proc1()
    On Error GoTo ErrH
    ' ... logic ...
    Exit Sub
ErrH:
    HandleError "Proc1"
End Sub

Function HandleError(source As String)
    MsgBox "Error in " & source & ": " & Err.Number & " - " & Err.Description
End Function
```

### 4. Long Parameter Lists
Functions with 5+ parameters.

**Fix:** Use parameter objects or Type structures.

```vba
' Before
Function CreateInvoice(custName As String, custAddr As String, _
    custCity As String, custZip As String, custPhone As String, _
    amount As Currency, discount As Single) As Long
    ...
End Function

' After
Type CustomerInfo
    Name As String
    Address As String
    City As String
    Zip As String
    Phone As String
End Type

Type InvoiceInfo
    Amount As Currency
    Discount As Single
End Type

Function CreateInvoice(customer As CustomerInfo, invoice As InvoiceInfo) As Long
    ...
End Function
```

### 5. Poor Error Handling
Missing or inadequate error handling.

**Fix:** Implement consistent error handling pattern.

```vba
' Standard pattern
Sub MyProcedure()
    On Error GoTo ErrHandler
    
    ' Main logic here
    
    Exit Sub
    
ErrHandler:
    ' Log error
    LogError "MyProcedure", Err.Number, Err.Description
    
    ' Show user-friendly message
    MsgBox "An error occurred. Please contact support.", vbCritical
    
    ' Cleanup if needed
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub
```

## Common Refactoring Patterns

### Extract Method
Break large procedures into smaller, focused ones.

```vba
' Before
Sub ProcessOrder()
    ' Validate customer
    If Len(customerName) = 0 Then
        MsgBox "Invalid customer"
        Exit Sub
    End If
    
    ' Calculate total
    Dim total As Currency
    total = qty * price
    If qty > 100 Then total = total * 0.9
    
    ' Save to database
    CurrentDb.Execute "INSERT INTO Orders..."
End Sub

' After
Sub ProcessOrder()
    If Not ValidateCustomer() Then Exit Sub
    
    Dim total As Currency
    total = CalculateOrderTotal(qty, price)
    
    SaveOrder total
End Sub

Private Function ValidateCustomer() As Boolean
    If Len(customerName) = 0 Then
        MsgBox "Invalid customer"
        ValidateCustomer = False
        Exit Function
    End If
    ValidateCustomer = True
End Function

Private Function CalculateOrderTotal(qty As Long, price As Currency) As Currency
    Dim total As Currency
    total = qty * price
    If qty > 100 Then total = total * 0.9
    CalculateOrderTotal = total
End Function

Private Sub SaveOrder(total As Currency)
    CurrentDb.Execute "INSERT INTO Orders..."
End Sub
```

### Replace Temp with Query
Eliminate temporary variables by using database queries.

```vba
' Before
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Customers WHERE ID = " & custID)
Dim custName As String
If Not rs.EOF Then
    custName = rs!CustomerName
End If
rs.Close

' After - Direct lookup function
Function GetCustomerName(custID As Long) As String
    GetCustomerName = Nz(DLookup("CustomerName", "Customers", "ID = " & custID), "")
End Function
```

### Introduce Parameter Object
Group related parameters into a structure.

```vba
' Before
Sub CreateReport(startDate As Date, endDate As Date, _
    includeDetails As Boolean, sortBy As String, _
    filterBy As String, groupBy As String)
    ...
End Sub

' After
Type ReportParameters
    StartDate As Date
    EndDate As Date
    IncludeDetails As Boolean
    SortBy As String
    FilterBy As String
    GroupBy As String
End Type

Sub CreateReport(params As ReportParameters)
    ...
End Sub
```

### Replace Conditional with Polymorphism
Use Select Case instead of nested If statements.

```vba
' Before
If type = "A" Then
    result = value * 1.1
ElseIf type = "B" Then
    result = value * 1.2
ElseIf type = "C" Then
    result = value * 1.3
Else
    result = value
End If

' After
Select Case type
    Case "A": result = value * 1.1
    Case "B": result = value * 1.2
    Case "C": result = value * 1.3
    Case Else: result = value
End Select
```

## Naming Conventions

### Hungarian Notation (Common in Access VBA)

**Controls:**
- `txt` - TextBox (txtCustomerName)
- `cbo` - ComboBox (cboCategory)
- `lst` - ListBox (lstProducts)
- `chk` - CheckBox (chkActive)
- `cmd` - CommandButton (cmdSave)
- `lbl` - Label (lblTitle)
- `frm` - Form (frmCustomers)
- `rpt` - Report (rptSales)

**Variables:**
- `str` - String (strName)
- `int` - Integer (intCount)
- `lng` - Long (lngID)
- `cur` - Currency (curTotal)
- `dbl` - Double (dblRate)
- `dat` - Date (datStart)
- `bln` - Boolean (blnActive)
- `var` - Variant (varResult)
- `rs` - Recordset (rsCustomers)
- `db` - Database (dbCurrent)

### Modern Naming (Recommended for refactored code)

Use clear, descriptive names without prefixes:

```vba
' Instead of
Dim strCustomerName As String
Dim intOrderCount As Integer

' Use
Dim customerName As String
Dim orderCount As Integer
```

## Performance Optimizations

### 1. Use Transactions for Batch Operations

```vba
Dim db As DAO.Database
Set db = CurrentDb

db.BeginTrans
On Error GoTo RollbackTrans

' Multiple inserts/updates
For i = 1 To 1000
    db.Execute "INSERT INTO..."
Next i

db.CommitTrans
Exit Sub

RollbackTrans:
    db.Rollback
    ' Handle error
End Sub
```

### 2. Avoid Repeated DLookup in Loops

```vba
' Bad
For i = 1 To 1000
    custName = DLookup("Name", "Customers", "ID = " & i)
    ' Process...
Next

' Good
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT ID, Name FROM Customers")
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Do While Not rs.EOF
    dict(rs!ID) = rs!Name
    rs.MoveNext
Loop
rs.Close

For i = 1 To 1000
    If dict.Exists(i) Then
        custName = dict(i)
        ' Process...
    End If
Next
```

### 3. Use SQL Instead of Recordset Loops

```vba
' Bad
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM Orders WHERE Status = 'Pending'")
Do While Not rs.EOF
    rs.Edit
    rs!Status = "Processed"
    rs.Update
    rs.MoveNext
Loop

' Good
CurrentDb.Execute "UPDATE Orders SET Status = 'Processed' WHERE Status = 'Pending'"
```

## Documentation Best Practices

### Function Headers

```vba
'===========================================================================
' FUNCTION: CalculateDiscount
' PURPOSE:  Calculates discount based on quantity and customer type
' PARAMETERS:
'   quantity (Long): Number of items ordered
'   customerType (String): "Regular", "Premium", or "VIP"
' RETURNS: Currency - Discount amount
' AUTHOR: [Name]
' DATE: 2026-01-28
' CHANGES:
'   2026-01-28 - Initial version
'===========================================================================
Function CalculateDiscount(quantity As Long, customerType As String) As Currency
    ...
End Function
```

### Inline Comments

```vba
' Good: Explain WHY, not WHAT
' Apply bulk discount for orders over 100 units
If quantity > 100 Then
    discount = basePrice * 0.1
End If

' Bad: Stating the obvious
' Check if quantity is greater than 100
If quantity > 100 Then
    discount = basePrice * 0.1
End If
```

## Testing Patterns

### Unit Test Template

```vba
'===========================================================================
' TEST MODULE: TestInvoiceCalculations
'===========================================================================
Sub TestCalculateDiscount()
    Debug.Print "Testing CalculateDiscount..."
    
    ' Test 1: Regular customer, small order
    Dim result As Currency
    result = CalculateDiscount(50, "Regular")
    Debug.Assert result = 0
    Debug.Print "  Test 1: PASS"
    
    ' Test 2: VIP customer, large order
    result = CalculateDiscount(200, "VIP")
    Debug.Assert result > 0
    Debug.Print "  Test 2: PASS"
    
    Debug.Print "All tests passed!"
End Sub
```

## Common Anti-Patterns to Avoid

1. **DoCmd.RunSQL without error handling**
2. **Global variables for everything**
3. **Forms referencing other forms directly (tight coupling)**
4. **Hard-coded connection strings**
5. **No Option Explicit**
6. **Using Goto except for error handling**
7. **Variants when specific types should be used**
8. **Not closing recordsets and database connections**

## Modernization Strategies

When refactoring legacy Access VBA:

1. **Add Option Explicit** to all modules
2. **Replace inline SQL** with saved queries or parameters
3. **Extract business logic** from form events
4. **Centralize configuration** (connection strings, file paths)
5. **Implement logging** for debugging
6. **Add error handling** consistently
7. **Document public interfaces**
8. **Create test procedures** for critical functions
