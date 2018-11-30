# VB Code Guidelines

<!-- TOC -->

- [VBA Code Guidelines](#vba-code-guidelines)
  - [General Advice](#general-advice)
  - [Parameters](#parameters)
  - [General errors](#general-errors)
  - [Variables](#variables)
    - [General](#general)
    - [Declaring](#declaring)
    - [Comments](#comments)
    - [Variants](#variants)
    - [Dates](#dates)
  - [General Naming Conventions](#general-naming-conventions)
    - [General](#general-1)
    - [Prefix](#prefix)
    - [Tag](#tag)
    - [Base name](#base-name)
    - [Qualifiers](#qualifiers)
    - [Arrays](#arrays)
    - [Constants](#constants)
  - [API Declaration](#api-declaration)
    - [Use unique alias names](#use-unique-alias-names)
  - [Form, Class & Module Naming](#form-class--module-naming)
    - [Internal Naming](#internal-naming)
    - [File naming](#file-naming)
    - [Object instance naming](#object-instance-naming)
    - [Notes](#notes)
  - [Naming Procedures/Functions/Parameters](#naming-proceduresfunctionsparameters)
  - [Function Names](#function-names)
    - [Function return values](#function-return-values)
    - [Parameters](#parameters-1)
  - [Naming Controls](#naming-controls)
    - [Introduction](#introduction)
    - [Control tags](#control-tags)
    - [Naming menu items](#naming-menu-items)
  - [**Naming Data Access Objects**](#naming-data-access-objects)
    - [ADO](#ado)
    - [ADO objects](#ado-objects)
    - [MS Access objects](#ms-access-objects)
  - [Layout](#layout)
    - [Indentation – tab width](#indentation--tab-width)
    - [Indentation - general](#indentation---general)
  - [Commenting Code](#commenting-code)
    - [Comments](#comments-1)
    - [Commenting code when doing maintenance work](#commenting-code-when-doing-maintenance-work)
    - [Etiquette when commenting code](#etiquette-when-commenting-code)
    - [Pre-compilation commands](#pre-compilation-commands)
  - [Error Handling](#error-handling)
    - [Generic error handler](#generic-error-handler)
    - [Error handling labels](#error-handling-labels)
 - [SQL Server stored procedures](#sql-store-procedures)
 	- [Overview](#overview)
 	- [Commenting Code ](#Commenting-Code )
 	- [Naming conventions](#Naming-conventions)
 - [Database Coding Standard and Guideline](#database-coding-standard-and-guideline)
   - [Naming](#naming)
   - [Structure](#structure)
   - [Formatting](#formatting)
   - [#Reference](#reference)
<!-- /TOC -->

## General Advice

- The first character of that name must be an alphabetic character
- Begin each separate word in a name with a capital letter, as in ```FindLastRecord``` and ```RedrawMyForm```.
- Begin function and method names with a verb, as in ```InitNameArray``` or ```CloseDialog```.
- Begin class, structure, module, and property names with a noun, as in ```EmployeeName``` or ```CarAccessory```.
- Begin interface names with the prefix "I", followed by a noun or a noun phrase, like IComponent, or with an adjective describing the interface's behavior, like IPersistable. Do not use the underscore, and use abbreviations sparingly, because abbreviations can cause confusion.
- Begin event handler names with a noun describing the type of event followed by the "EventHandler" suffix, as in "MouseEventHandler".
- In names of event argument classes, include the "EventArgs" suffix.
- If an event has a concept of "before" or "after," use a suffix in present or past tense, as in "ControlAdd" or "ControlAdded".
- For long or frequently used terms, use abbreviations to keep name lengths reasonable, for example, "HTML", instead of "Hypertext Markup Language". In general, variable names greater than 32 characters are difficult to read on a monitor set to a low resolution. Also, make sure your abbreviations are consistent throughout the entire application. Randomly switching in a project between "HTML" and "Hypertext Markup Language" can lead to confusion.
- Avoid using names in an inner scope that are the same as names in an outer scope. Errors can result if the wrong variable is accessed. If a conflict occurs between a variable and the keyword of the same name, you must identify the keyword by preceding it with the appropriate type library. For example, if you have a variable called Date, you can use the intrinsic Date function only by calling DateTime.Date.

## Parameters

Avoid confusion over ByVal and ByRef. Be aware of the default for
parameters being ByRef. Be explicit when passing parameters.

- [Force] Only use ByRef where you intend to modify the parameter and
  pass the change back to the Caller.
- [Force] Pass parameter ByVal if they are not to be changes
- [Suggest] Explicitly use ByRef in an input parameter is to be
  changed, but watch for signs to redesign.

Pass by Reference example:

```vba
Private Sub ChangeRefValue()
    Dim intX As Integer
    intX = 1
    Debug.Print intX    ' print value of intX is 1

    Call ChangeValueByRef(intX)

    Debug.Print intX    ' print value of intX is 6
End Sub

Sub ChangeValueByRef(ByRef intY As Integer)
    intY = intY + 5
End Sub
```

Pass by Value example:

```vba
Public Sub Load(ByVal strName As String, ByVal strPhone As String)
```

## General errors

Error handling must be used wherever practicable i.e. within each
procedure. Use On Error Goto ErrHandler Handle errors where they
occur. This may involve handling the error and raising it to the
client code.

## Variables

### General

[Force] Where global variables are used, they must all be defined in one
module.

[Force] Hard-coded(Magic) numbers and strings should be made into constants.

[Force] Use explicit data casting fcuntion `Cstr()`, `CDate()`, `Cbool()` etc.

### Declaring

[Force] Variables must be dimensioned on separate lines, and should specify a
datatype (except where this is not possible – as when using certain
scripting languages).

### Comments

[Suggest] All variables must be declared at the top of each procedure or module
and must ideally be grouped so that all variable types are placed
together.

### Variants

[Suggest] Variants may be used where appropriate (e.g. to hold arrays returned
by a function, or where Nulls may be encountered), but alternative
data types should be used where possible.

[Suggest] Advoid using Variants data type unless absolutely necessary.
Variants are slower then native types, when there's large dataset
or a big macro project.

### Dates

[Force] Where dates are displayed to users you should avoid ambiguous formats
where either years or days vs. months might be confused (such as
DD/MM/YY), however the ultimate decision maker on this issue is the
customer.

Where dates are being handled “behind the scenes” care should be taken
to avoid UK/US format confusion. Particular care should be taken when
including UK-format dates in literal SQL strings (where the target
Microsoft application may expect dates to be in US format). Where
there is the slightest possibility of doubt pass the year, month and
day parts separately into DateSerial, of format them in the
universally acceptable ISO format YYYY-MM-DD.

## General Naming Conventions

### General

Object names are made up of four parts: prefix tag base name qualifier
The four parts are assembled as follows:
[prefixes]tag[BaseName][qualifier] Note: The brackets denote that
these components are optional and are not part of the name.

### Prefix

Prefixes and tags are always lowercase so your eye goes past them to
the first uppercase letter where the base name begins. This makes the
names more readable. The base and qualifier components begin with an
uppercase letter.

| Prefix | Use                | Notes                             |
| ------ | ------------------ | --------------------------------- |
| None   | Local to procedure | No scope prefix as in: dblMaximum |
| m\_    | Module level scope | m_strPolicyHolder                 |
| g\_    | Global scope       | g_intCarsLast                     |

### Tag

The tag is the only required component, but in almost all cases the
name will have the base name component since you need to be able to
distinguish two objects of the same type.

| Variable type     | Tag | Notes         |
| ----------------- | --- | ------------- |
| Boolean           | bln | blnFound      |
| Byte              | byt | bytRasterData |
| Currency          | cur | curRevenue    |
| Date (Time)       | dat | datStart      |
| Double            | dbl | dblTolerance  |
| Enum              | enm | enmColours    |
| Integer           | int | intQuantity   |
| Long              | lng | lngDistance   |
| Single            | sng | sngAverage    |
| String            | str | strFName      |
| User-defined type | udt | udtEmployee   |
| Variant           | var | varCheckSum   |

[suggest] To avoid defining an ambiguous variable, it is strongly suggest to use
3-letter abbreviations instead of using a single-letter abbreviations

Positive Example:
```vba
Dim intProductID As Integer
```

Negative Example:
```vba
Dim iProductID As Integer   ' Too short abbreviative tag definition
Dim orderID As Integer      ' Meaningless definition
```

### Base name

The base name succinctly describes the object, not its class. That is,
a base name for a variable for an invoice form must be InvoiceEntry
not InvoiceForm as the tag will describe the object. The base name is
composed in the form Noun[Verb]. For example, in the variable name
blnColourMatch "ColourMatch" is the base name. Naming variables in
this way helps to keep them grouped together in documentation and
cross-referencing because they will be sorted together alphabetically.

### Qualifiers

Object qualifiers may follow a name and further clarify names that are
similar. Continuing with our previous example, if you kept two indexes
to an array, one for the first item and one for the last this entails
two qualified variables such as intColourMatchFirst and
intColourMatchLast. Other examples of qualifiers:

| Usage                   | Qualifier | Example         |
| ----------------------- | --------- | --------------- |
| Current element of set  | Cur       | intCarsCur      |
| First element of set    | First     | intCarsFirst    |
| Last element of set     | Last      | intCarsLast     |
| Next element of set     | Next      | strCustomerNext |
| Previous element of set | Prev      | strCustomerPrev |
| Lower limit of range    | Min       | strNameMin      |
| Upper limit of range    | Max       | strNameMax      |
| Source                  | Src       | lngBufferSrc    |
| Destination             | Dest      | lngBufferDest   |

### Arrays

[Force] Array names must be prefixed with "a" or "arr". The upper and lower
bounds ofthe array must be declared explicitly (unless they’re not known at
design-time).

Positive Example:
```vba
Dim astrMonths(1 To 12) As String
```

Negative Example:
```vba
Dim strMonths(1 To 12) As String
```

### Constants

Each word must be capitalised and the words separated with an
underscore. The base name must be a description of what the constant
represents.

Example:

```vba
User defined constant: g_intERR_INVALID_NAME
Visual Basic: vbArrowHourglass
```

## API Declaration

API declarations must be laid out so that they are easily readable on
the screen.

```vba
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
```

### Use unique alias names

In VB you can call external procedures in DLLs when you know the entry
point (the name of the function in the DLL). However, the caveat is
that you can only declare the external procedure once. If you load a
library that calls the same Windows API that your module calls, you
will get the infamous error, “Tried to load module with duplicate
procedure definition.”

```
Declare smg_GetActiveWindow Lib "Kernel" Alias _
	"GetActiveWindo" () As Integer
```

## Form, Class & Module Naming

### Internal Naming

(i.e. the name assigned to the module within the VB Properties)

| Module Type      | Prefix | Example      |
| ---------------- | ------ | ------------ |
| Form             | frm    | frmLogon     |
| Standard module  | mod    | modUtilities |
| Class module     | C      | CPerson      |
| Collection class | C      | CPersons1    |
| Interface class  | I      | IPerson      |

### File naming

(i.e. the name assigned to the module when saving the physical file)

| Module Type      | Prefix | Example          |
| ---------------- | ------ | ---------------- |
| Form             | frm    | frmLogon.frm     |
| Standard module  | mod    | modUtilities.bas |
| Class module     | C      | CPerson.cls      |
| Collection class | C      | CPersons.cls1    |
| Interface class  | I      | IPerson.cls      |

### Object instance naming

(i.e. the name assigned when declaring a variable based on the form or
class)

| Instance of | Prefix | Example    |
| ----------- | ------ | ---------- |
| Form        | frm    | frmLogon   |
| Class       | obj    | objPerson  |
| Collection  | obj    | objPersons |

### Notes

1 Classes which hold collections should have the same “C” prefix as
any other classes, but should have a plural name (based on the type of
objects held in the collection. E.g. a class to hold a single person
would be named CPerson, whereas a collection of Person objects would
be named CPersons.

## Naming Procedures/Functions/Parameters

## Function Names

Tags should not be prefixed to Function or Sub names, but **should**
be appended to the parameters of these routines. For example:

_**Correct approach for internal function:**_

```vba
Private Function TotalUp(ByVal sngSubTotal As Single) As Integer
```

### Function return values

Function return values should usually be held in a temporary variable
and then assigned to the function variable at the end of the routine.
This has two benefits. The code is not specific to the name of the
function so portability is aided when cutting and pasting part of the
function code elsewhere; also the value of the function variable may
be used in calculations, otherwise a recursive call would be
generated. Example:

```vba
Private Function Example(ByVal argintA as Integer) as Single
    Dim sngRetVal as Single

    ' Set default value
    sngRetVal = 0

    <code block>

    ' Set the Function value
    Example = sngRetVal
End Function
```

### Parameters

Should you find it useful, you may also prefix parameter names with
arg to avoid confusion between variables passed as parameters and
those local to the subroutine . Example:

```vba
Private Function DoSomething(ByVal argstrMessage as String) as String
```

However, should you choose to adopt this standard it must be applied
consistently across the entire project

## Naming Controls

### Introduction

Controls must be named with uniform prefixes strictly adhering to the
following list.

### Control tags

| Object Type                   | Tag  | Notes           |
| ----------------------------- | ---- | --------------- |
| Check box                     | chk  | chkReadOnly     |
| Combo box, drop-down list box | cbo  | cboEnglish      |
| Command button                | cmd  | cmdExit         |
| Common dialog                 | dlg  | dlgFileOpen     |
| Control                       | ctl  | ctlCurrent      |
| Form                          | frm  | frmEntry        |
| Frame                         | fra  | fraLanguage     |
| Grid                          | grd  | grdPrices       |
| Image                         | img  | imgIcon         |
| Key status                    | key  | keyCaps         |
| Label                         | lbl  | lblHelpMessage  |
| Line                          | lin  | linVertical     |
| List box                      | lst  | lstPolicyCodes  |
| Menu                          | mnu  | mnuFileOpen     |
| Report                        | rpt  | rptQtr1Earnings |
| Shape                         | shp  | shpCircle       |
| Text box                      | txt  | txtLastName     |
| True DBGrid                   | tdbg | tdbgRecords     |
| Timer                         | tmr  | tmrAlarm        |
| ImageList                     | ils  | ilsAllIcons     |
| Toolbar                       | tlb  | tlbActions      |
| TabStrip                      | tab  | tabOptions      |
| ListView                      | lvw  | lvwHeadings     |

### Naming menu items

The number of menu options can be great, so it is recommended that
there be a standard for the names given to menus. The tag for any menu
item whether an option or title is mnu. Prefixing must ideally
continue beyond the initial prefix. The first prefix after mnu is the
menu bar title followed by the option then any subsequent option.

Example:

```
Top level menu item – mnuFile
Menu sub item – mnuFileSave
```

## **Naming Data Access Objects**

### ADO

If you include references to both ADO and DAO in the same project you
must explicitly specify which object model you wish to use when
declaring variables. Example:

```vba
Dim cnnStore As ADODB.Connection
Dim cnnOther As DAO.Connection
```

### ADO objects

| Object Type | Tag | Example     |
| ----------- | --- | ----------- |
| Command     | cmd | cmdBooks    |
| Connection  | cnn | cnnLibrary  |
| Parameter   | prm | prmTitle    |
| Error       | err | errLoop     |
| Recordset   | rst | rstForecast |

### MS Access objects

The following is a suggested naming convention for use with MS Access
objects – you may find it useful for larger Access projects which have
many objects within the same database.

| Object Type        | Tag  | Example           |
| ------------------ | ---- | ----------------- |
| Table              | tbl  | tblCustomer       |
| Query (select)     | qry  | qryOverAchiever   |
| Query (append)     | qapp | qappNewProduct    |
| Query (crosstab)   | qxtb | qxtbRegionSales   |
| Query (delete)     | qdel | qdelOldAccount    |
| Query (make table) | qmak | qmakShipTo        |
| Query (update)     | qupd | qupdDiscount      |
| Form               | frm  | frmCustomer       |
| Form (dialog)      | fdlg | fdlgLogin         |
| Form (message)     | fmsg | fmsgWait          |
| Form (subform)     | fsub | fsubOrder         |
| Report             | rpt  | rptInsuranceValue |
| Report (subreport) | rsub | rsubOrder         |
| Macro (menu)       | mmnu | mmnuEntryFormFile |
| Module             | mod  | modBilling        |

## Layout

### Indentation – tab width

When working in a VB or VBA design environment you **must** have the
**Tab Width** set to 4 (see the Editor tab in Tools > Options). This
is the default VB setting, and using it ensures compatibility when
code is worked on by more than one person.

### Indentation - general

Code must be indented consistently adhering to the following rules:

- Declarations must not be indented.
- On Error statements and line labels/numbers must not be indented.
- Start code indented to one tab stop.
- Code within If-Else-EndIf, For-Next, Do While/Until and any other
  loops must be indented a further tab stop within the body.
- Code between add/edit and update statements must be indented a
  further tab stop.
- Case statements must be indented to one stop after the Select Case.
  Code following the Case statements must be indented a further Tab
  stop.
- Code between With and End With statements must be indented by one
  tab stop.
- Code within error trap must be indented by to one tab stop. Example

```vba
Dim strTest as String
Dim wrk as Workspace
On Error Goto ErrHandler
    If strTest = "" Then
        strTest = "Nothing"
    Else
        strTest = ""
    EndIf

    Do While Not rst.EOF
        rst.Add
        rst(0) = strTest
        rst.Update
    Loop

    Select Case strTest
        Case ""
            <code block>
        Case Else
            <code block>
        End Select
ExitHere:
    Exit Sub
ErrHandler:
    Resume ExitHere
```

## Commenting Code

### Comments

Remember the following points:

- Code must be commented appropriately. The goal should be to improve
  understanding and maintainability of the code.
- Comments should explain the reasoning behind the code. It may be
  obvious to the original developer what a piece of code does but
  somebody reading it may have no idea why it has to be there. When
  you write a piece of code, imagine someone else having to read
  through it 3 months later. Will it make sense to them?
- Important variable declarations may include an inline comment
  describing the use of the variable being declared.

Example:

```vba
Dim strLookUp as String 'Accepts value from user to search for
```

- Comments for individual lines appear above, or of the code to which
  they refer.
- The functional overview comment of a procedure may be indented one
  space to aid readability.

Example:

```vba
Public Sub DeleteCustomer(ByVal argintID As Long)
    'Removes customer from Database
    cnVideo.Execute "DELETE FROM Customer WHERE CustomerID=" & argintID
End Sub
```

### Commenting code when doing maintenance work

Avoid over-commenting code when doing maintenance work. Bear in mind
the need to maintain overall clarity in the code, and remember that
revision history should be taken care of by SourceSafe Make sure that
any existing comments still make sense **after** you’ve made your
changes - paying particular attention to any comments/explanations in
the header of the routine. You are responsible for ensuring that
**all** existing comments remain accurate (and that they still make
sense) after your changes have been implemented. Although SourceSafe
controls the history, It is handy to future users if new blocks of
code are commented with the date, initials of developer and the CR
number to aim future developers reading the code.

### Etiquette when commenting code

When you include one or more routines written by other developers in
your project you should ensure that any author (and
assumption/purpose) information in the routine header is kept
accurate. You should probably retain the original author’s name, but
you **must** also include your own name if you have changed it in any
way at all.

### Pre-compilation commands

These are treated as a code IF statement would be. All code relating
to the condition must be indented as if it was a normal IF block.
These can be useful for including/excluding debug code etc. For
example:

```vba
#Const DebugMode = True
#IF  DebugMode THEN
    <code block>
#ELSE
    <code block>
#ENDIF
```

## Error Handling

### Generic error handler

Consistent error handlers must be implemented. The following error
handler should be used:

```vba
On Error GoTo ErrHandler
    <code block>

ExitHere:
On Error Resume Next
    <code block>
Exit Sub
ErrHandler:
    [WriteErrLog Err.Number]
        Select Case Err.Number
            Case <Err No>
                Resume Next
            Case <Err No>
                Resume ExitHere
            Case Else
                ' Unexpected Error
                Resume ExitHere
        End Select
End Sub
```

### Error handling labels

The labels **ErrHandler** and **ExitHere** are used both for
consistency across routines, and to facilitate easier copying and
pasting of error handlers between routines.

## SQL Server stored procedures
### Overview
A stored procedure is nothing more than prepared SQL code that you save so you can reuse the code over and over again.  So if you think about a query that you write over and over again, instead of having to write that query each time you would save it as a stored procedure and then just call the stored procedure to execute the SQL code that you saved as part of the stored procedure.

In addition to running the same SQL code over and over again you also have the ability to pass parameters to the stored procedure, so depending on what the need is the stored procedure can act accordingly based on the parameter values that were passed.

Take a look through each of these topics to learn how to get started with stored procedure development for SQL Server.

You can either use the outline on the left or click on the arrows to the right or below to scroll through each of these topics.
### Commenting Code 
SQL Server offers two types of comments in a stored procedure; line comments and block comments.   The following examples show you how to add comments using both techniques.  Comments are displayed in green in a SQL Server query window.

Line Comments
To create line comments you just use two dashes "--" in front of the code you want to comment.  You can comment out one or multiple lines with this technique.

In this example the entire line is commented out.
```
-- this procedure gets a list of addresses based 
-- on the city value that is passed
CREATE PROCEDURE dbo.uspGetAddress @City nvarchar(30)
AS
SELECT * 
FROM Person.Address
WHERE City = @City
GO
```
This next example shows you how to put the comment on the same line.
```
-- this procedure gets a list of addresses based on the city value that is passed
CREATE PROCEDURE dbo.uspGetAddress @City nvarchar(30)
AS
SELECT * 
FROM Person.Address
WHERE City = @City -- the @City parameter value will narrow the search criteria
GO
```
Block Comments
To create block comments the block is started with "/*" and ends with "*/".   Anything within that block will be a comment section.
```
/* 
-this procedure gets a list of addresses based 
 on the city value that is passed
-this procedure is used by the HR system      
*/
CREATE PROCEDURE dbo.uspGetAddress @City nvarchar(30)
AS
SELECT * 
FROM Person.Address
WHERE City = @City
GO
```
Combining Line and Block Comments
You can also use both types of comments within a stored procedure.
```
/* 
-this procedure gets a list of addresses based 
 on the city value that is passed
-this procedure is used by the HR system      
*/
CREATE PROCEDURE dbo.uspGetAddress @City nvarchar(30)
AS
SELECT * 
FROM Person.Address
WHERE City = @City -- the @City parameter value will narrow the search criteria
GO
```
### Naming conventions

#### Naming Stored Procedure Action
I liked to first give the action that the stored procedure takes and then give it a name representing the object it will affect.

So based on the actions that you may take with a stored procedure, you may use:
```
Insert
Delete
Update
Select
Get
Validate
etc...
```
So here are a few examples:
```
uspInsertPerson
uspGetPerson
spValidatePerson
SelectPerson
etc...
```
Another option is to put the object name first and the action second, this way all of the stored procedures for an object will be together.
```
uspPersonInsert
uspPersonDelete
uspPersonGet
etc...
```
Again, this does not really matter what action words that you use, but this will be helpful to classify the behavior characteristics.
#### Naming Stored Procedure Object
The last part of this is the object that you are working with.  Some of these may be real objects like tables, but others may be business processes.  Keep the names simple, but meaningful.  As your database grows and you add more and more objects you will be glad that you created some standards.

So some of these may be:
```
uspInsertPerson - insert a new person record
uspGetAccountBalance - get the balance of an account
uspGetOrderHistory - return list of orders
```
#### Schema Names
Another thing to consider is the schema that you will use when saving the objects.  A schema is the a collection of objects, so basically just a container.  This is useful if you want to keep all utility like objects together or have some objects that are HR related, etc...

This logical grouping will help you differentiate the objects further and allow you to focus on a group of objects.

Here is a simple example to create a new schema called "HR" and giving authorization to this schema to "DBO".
```
CREATE SCHEMA [HumanResources] AUTHORIZATION [dbo]
```
#### Putting It All Together
So you basically have four parts that you should consider when you come up with a naming convention:
```
Schema
Prefix
Action
Object
```
Take the time to think through what makes the most sense and try to stick to your conventions.
## **Database Coding Standard and Guideline**
 ### Naming

**Tables**: Rules: Pascal notation; end with an ‘s’
- Examples: Products, Customers
- Group related table names(1)

**Stored Procs**: Rules: spAppName_GroupNameAction
- Examples: spOrders_GetNewOrders, spProducts_UpdateProduct

**Triggers**: Rules: TR_TableName_action
- Examples: TR_Orders_UpdateProducts
- Notes: The use of triggers is discouraged

**Indexes**: Rules: IX_TableName_columns separated by "-"
- Examples: IX_Products_ProductID

**Primary Keys**: Rules: PK_TableName
- Examples: PK_Products

**Foreign Keys**: Rules: FK_TableName1_TableName
- Example: FK_Products_Orderss

**Defaults: Rules**: DF_TableName_ColumnName
- Example: DF_Products_Quantity

**Columns**: If a column references another table’s column, name it table nameID
- Example: The Customers table has an ID column
- The Orders table should have a CustomerID column

**General Rules:**

- Do not use spaces in the name of database objects
- Do not use SQL keywords as the name of database objects
- In cases where this is necessary, surround the object name with brackets, such as [Year]
- Do not prefix stored procedures with ‘sp_’(2)
- Prefix table names with the owner name  (3)

### Structure

- Each table must have a primary key
 - In most cases it should be an IDENTITY column named ID
- Normalize data to third normal form
 - Do not compromise on performance to reach third normal form. Sometimes, a little de-normalization results in better performance.
- Do not use TEXT as a data type; use the maximum allowed characters of VARCHAR instead
- In VARCHAR data columns, do not default to NULL; use an empty string instead
- Columns with default values should not allow NULLs
- As much as possible, create stored procedures on the same database as the main tables they will be accessing

### Formatting

- Use upper case for all SQL keywords
 - SELECT, INSERT, UPDATE, WHERE, AND, OR, LIKE, etc.
- Indent code to improve readability
- Comment code blocks that are not easily understandable
 - Use single-line comment markers(?)
 - Reserve multi-line comments (/*.. ..*/) for blocking out sections of code
- Use single quote characters to delimit strings.
  - Nest single quotes to express a single quote or apostrophe within a string
   - For example, SET @sExample = ‘SQL”s Authority’
- Use parentheses to increase readability
 - WHERE (color=’red’ AND (size = 1 OR size = 2))
- Use BEGIN..END blocks only when multiple statements are present within a conditional code segment.
- Use one blank line to separate code sections.
- Use spaces so that expressions read like sentences.
 - fillfactor = 25, not fillfactor=25
- Format JOIN operations using indents
 - Also, use ANSI Joins instead of old style joins (4)
- Place SET statements before any executing code in the procedure.
- Optimize queries using the tools provided by SQL Server(5)
- Do not use SELECT *
- Return multiple result sets from one stored procedure to avoid trips from the application server to SQL server
- Avoid unnecessary use of temporary tables
 - Use ‘Derived tables’ or CTE (Common Table Expressions) wherever possible, as they perform better (6)
- Avoid using <> as a comparison operator
 - Use ID IN(1,3,4,5) instead of ID <> 2
- Use SET NOCOUNT ON at the beginning of stored procedures (7)
- Do not use cursors or application loops to do inserts (8)
 - Instead, use INSERT INTO
- Fully qualify tables and column names in JOINs
- Fully qualify all stored procedure and table references in stored procedures.
- Do not define default values for parameters.
 - If a default is needed, the front end will supply the value.
- Do not use the RECOMPILE option for stored procedures.
- Place all DECLARE statements before any other code in the procedure.
- Do not use column numbers in the ORDER BY clause.
- Do not use GOTO.
- Check the global variable @@ERROR immediately after executing a data manipulation statement (like INSERT/UPDATE/DELETE), so that you can rollback the transaction if an error occurs
 - Or use TRY/CATCH
- Do basic validations in the front-end itself during data entry
- Off-load tasks, like string manipulations, concatenations, row numbering, case conversions, type conversions etc., to the front-end applications if these operations are going to consume more CPU cycles on the database server
- Always use a column list in your INSERT statements.
 - This helps avoid problems when the table structure changes (like adding or dropping a column).
- Minimize the use of NULLs, as they often confuse front-end applications, unless the applications are coded intelligently to eliminate NULLs or convert the NULLs into some other form.
 - Any expression that deals with NULL results in a NULL output.
 - The ISNULL and COALESCE functions are helpful in dealing with NULL values.
- Do not use the identitycol or rowguidcol.
- Avoid the use of cross joins, if possible.
- When executing an UPDATE or DELETE statement, use the primary key in the WHERE condition, if possible. This reduces error possibilities.
- Avoid using TEXT or NTEXT datatypes for storing large textual data. (9)
 - Use the maximum allowed characters of VARCHAR instead
- Avoid dynamic SQL statements as much as possible. (10)
- Access tables in the same order in your stored procedures and triggers consistently. (11)
- Do not call functions repeatedly within your stored procedures, triggers, functions and batches. (12)
- Default constraints must be defined at the column level.
- Avoid wild-card characters at the beginning of a word while searching using the LIKE keyword, as these results in an index scan, which defeats the purpose of an index.
- Define all constraints, other than defaults, at the table level.
- When a result set is not needed, use syntax that does not return a result set. (13)
- Avoid rules, database level defaults that must be bound or user-defined data types. While these are legitimate database constructs, opt for constraints and column defaults to hold the database consistent for development and conversion coding.
- Constraints that apply to more than one column must be defined at the table level.
- Use the CHAR data type for a column only when the column is non-nullable. (14)
- Do not use white space in identifiers.
- The RETURN statement is meant for returning the execution status only, but not data.

### # **Reference**:
1) Group related table names:

Products_USA
Products_India
Products_Mexico

2) The prefix sp_ is reserved for system stored procedures that ship with SQL Server. Whenever SQL Server encounters a procedure name starting with sp_, it first tries to locate the procedure in the master database, then it looks for any qualifiers (database, owner) provided, then it tries dbo as the owner. Time spent locating the stored procedure can be saved by avoiding the “sp_” prefix.

3) This improves readability and avoids unnecessary confusion. Microsoft SQL Server Books Online states that qualifying table names with owner names helps in execution plan reuse, further boosting performance.

4)
False code:
> SELECT *
FROM Table1, Table2
WHERE Table1.d = Table2.c

True code:
> SELECT *
FROM Table1
INNER JOIN Table2 ON Table1.d = Table2.c

5) Use the graphical execution plan in Query Analyzer or SHOWPLAN_TEXT or SHOWPLAN_ALL commands to analyze your queries. Make sure your queries do an “Index seek” instead of an “Index scan” or a “Table scan.” A table scan or an index scan is a highly undesirable and should be avoided where possible.

6) Consider the following query to find the second highest offer price from the Items table:
> SELECT MAX(Price)
FROM Products
WHERE ID IN
(
SELECT TOP 2 ID
FROM Products
ORDER BY Price DESC
)

The same query can be re-written using a derived table, as shown below, and it performs generally twice as fast as the above query:

> SELECT MAX(Price)
FROM
(
SELECT TOP 2 Price
FROM Products
ORDER BY Price DESC
)

7) This suppresses messages like ‘(1 row(s) affected)’ after executing INSERT, UPDATE, DELETE and SELECT statements. Performance is improved due to the reduction of network traffic.

8) Try to avoid server side cursors as much as possible. Always stick to a ‘set-based approach’ instead of a ‘procedural approach’ for accessing and manipulating data. Cursors can often be avoided by using SELECT statements instead. If a cursor is unavoidable, use a WHILE loop instead. For a WHILE loop to replace a cursor, however, you need a column (primary key or unique key) to identify each row uniquely.

9) You cannot directly write or update text data using the INSERT or UPDATE statements. Instead, you have to use special statements like READTEXT, WRITETEXT and UPDATETEXT. So, if you don’t have to store more than 8KB of text, use the CHAR(8000) or VARCHAR(8000) datatype instead.

10) Dynamic SQL tends to be slower than static SQL, as SQL Server must generate an execution plan at runtime. IF and CASE statements come in handy to avoid dynamic SQL.

11) This helps to avoid deadlocks. Other things to keep in mind to avoid deadlocks are:

Keep transactions as short as possible.
Touch the minimum amount of data possible during a transaction.
Never wait for user input in the middle of a transaction.
Do not use higher level locking hints or restrictive isolation levels unless they are absolutely needed.
12) You might need the length of a string variable in many places of your procedure, but don’t call the LEN function whenever it’s needed. Instead, call the LEN function once and store the result in a variable for later use.

13)

> IF EXISTS (
 SELECT 1
 FROM Products
 WHERE ID = 50)
 
Instead Of:

> IF EXISTS (
 SELECT COUNT(ID)
 FROM Products
 WHERE ID = 50)
 
14) CHAR(100), when NULL, will consume 100 bytes, resulting in space wastage. Preferably, use VARCHAR(100) in this situation. Variable-length columns have very little processing overhead compared with fixed-length columns
