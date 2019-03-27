Option Explicit

'Variables to connect to databases
Public MstrDb As New ADODB.Connection
Public FndStsDb As New ADODB.Connection
Public SxbtDb As New ADODB.Connection

'Variable for opening recordsets in database objects
Public rst As New ADODB.Recordset

'Variables for Formatting Fitments
Public Reopen As Boolean

'Variables for SourceForm
Public PartName As String
Public PartTypeVar As String
Public Brand As String
Public InterchangeSource As String
Public FitmentSource As String
Public Formatted As Boolean

'Variable for SKUForm
Public SKU As String

'Variables for ListAmazon form
Public PageCount As Integer
Public ListType As String
Public TemplateType As String
Public TemplateVersion As String
Public TemplateSignature As String
Public TemplateAmazonUse As String
Public LabelRow As Integer
Public NameRow As Integer

'Variables for Finale Product
Public FinaleFields() As Variant
Public CategoriesArr() As Variant

'Variables for Inventory
Public InventoryType As String

'For debugging
Public BypassRibbon As Boolean
