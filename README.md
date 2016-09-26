# VBA Code Guidelines

## General Advice

Always use Option Explicit as the first line in every code module. To switch this on automatically check Require Variable Declaration in Tools>Options>Editor.

## Parameters

Avoid confusion over ByVal and ByRef. Be aware of the default for parameters being ByRef. Be explicit when passing parameters.
 
Example:
```
Public Sub Load(ByVal strName As String, ByVal strPhone As String)
```

## General errors

Error handling must be used wherever practicable i.e. within each procedure.
Use On Error Goto ErrHandler
Handle errors where they occur. This may involve handling the error and raising it to the client code.

# Variables

## General
Where global variables are used, they must all be defined in one module.

## Declaring

Variables must be dimensioned on separate lines, and should specify a datatype (except where this is not possible – as when using certain scripting languages).

## Comments

All variables must be declared at the top of each procedure or module and must ideally be grouped so that all variable types are placed together.

## Variants

Variants may be used where appropriate (e.g. to hold arrays returned by a function, or where Nulls may be encountered), but alternative data types should be used where possible.

## Dates

Where dates are displayed to users you should avoid ambiguous formats where either years or days vs. months might be confused (such as DD/MM/YY), however the ultimate decision maker on this issue is the customer.

Where dates are being handled “behind the scenes” care should be taken to avoid UK/US format confusion.  Particular care should be taken when including UK-format dates in literal SQL strings (where the target Microsoft application may expect dates to be in US format).  Where there is the slightest possibility of doubt pass the year, month and day parts separately into DateSerial, of format them in the universally acceptable ISO format YYYY-MM-DD. 

# General Naming Conventions

## General

Object names are made up of four parts: 
prefix
tag
base name
qualifier
The four parts are assembled as follows:	
[prefixes]tag[BaseName][Qualifier]
Note: The brackets denote that these components are optional and are not part of the name.

## Prefix

Prefixes and tags are always lowercase so your eye goes past them to the first uppercase letter where the base name begins.  This makes the names more readable.  The base and qualifier components begin with an uppercase letter.

| Prefix | Use | Notes |
| --- | --- | --- |
| None | Local to procedure | No scope prefix as in: dblMaximum |
| m_ | Module level scope | m_strPolicyHolder |
| g_ | Global scope | g_intCarsLast |

## Tag

The tag is the only required component, but in almost all cases the name will have the base name component since you need to be able to distinguish two objects of the same type.

| Variable type | Tag | Notes |
| --- | --- | --- |
| Boolean | bln | blnFound |
| Byte | byt | bytRasterData |
| Currency | cur | curRevenue |
