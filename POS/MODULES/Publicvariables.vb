Imports MySql.Data.MySqlClient
Imports System.Drawing.Printing
Module publicVariables
    'Connection
    Public localconn As MySqlConnection
    Public cloudconn As MySqlConnection
    Public cmd As MySqlCommand
    Public cloudcmd As New MySqlCommand
    Public da As MySqlDataAdapter
    Public dr As MySqlDataReader
    Public connectionstring As String
    '=============================================================================================================
    'Data Table
    Public dt As DataTable
    Public sql As String
    Public param As MySqlParameter
    Public ds As DataSet
    '
    Public DisableFormClose As Boolean = False
    '
    'Public ProductDTUpdate As DataTable = New DataTable
    'Public InventoryDTUpdate As DataTable = New DataTable
    'Public FormulaDTUpdate As DataTable = New DataTable
    'Public CategoryDTUpdate As DataTable = New DataTable
    Public Messageboolean As Boolean = False
    '=============================================================================================================
    'POS
    Public SINumber As Integer

    Public SiNumberToString As String
    Public TRANSACTIONMODE As String = "Walk-In"
    Public SeniorPWd As Decimal
    Public SeniorPWdDrinks As Decimal
    Public ThisIsMyInventoryID
    Public Shift As String
    Public BeginningBalance As Decimal
    Public EndingBalance As Decimal
    Public payment As Boolean = False
    Public posandpendingenter As Boolean = False
    Public productprice As Integer
    Public Deleteitem As Boolean = False

    Public qtyisgreaterthanstock As Boolean = False
    Public hastextboxqty As Boolean = False
    Public productID
    Public getmunicipality
    Public getprovince
    Public Couponisavailable As Boolean
    Public DiscountType As String
    Public modeoftransaction As Boolean
    Public SyncIsOnProcess As Boolean

    Public CouponApplied As Boolean = False
    Public CouponName As String
    Public CouponDesc As String
    Public CouponLine As Integer = 10
    Public CouponTotal As Decimal
    Public SeniorGCDiscount As Boolean = False
    '=============================================================================================================
    'POS INFORMATION
    Public ClientCrewID As String
    Public ClientRole As String
    Public ClientGuid As String
    Public ClientStoreID As String
    Public ClientBrand As String
    Public ClientLocation As String
    Public ClientPostalCode As String
    Public ClientAddress As String
    Public ClientBrgy As String
    Public ClientMunicipality As String
    Public ClientProvince As String
    Public ClientTin As String
    Public ClientTel As String
    Public ClientStorename As String
    Public ClientProductKey As String
    Public ClientMIN As String
    Public ClientMSN As String
    Public ClientPTUN As String

    '==CONNECTION STRINGS
    Public LocalConnectionString As String
    Public CloudConnectionString As String
    Public ValidLocalConnection As Boolean = False
    Public ValidCloudConnection As Boolean = False
    '==SETTINGS
    Public S_SIBeg As Integer
    Public S_ExportPath As String
    Public S_Tax As String
    Public S_SIFormat As String
    Public S_Terminal_No As String
    Public S_ZeroRated As String
    Public S_Zreading
    Public S_Batter As String
    Public S_Brownie_Mix As String
    Public S_Upgrade_Price As String
    Public S_Backup_Interval As String
    Public S_Backup_Date As String
    Public S_Logo As String
    Public S_Icon As String
    Public S_Layout As String
    Public S_Print As String
    Public S_Reprint As String
    Public S_Print_XZRead As String
    Public S_Print_Returns As String
    Public S_Waffle_Bag As Integer
    Public S_Packets As Integer
    Public S_PrintCount As Integer
    Public S_Dev_Comp_Name As String
    'Training MOde
    Public S_TrainingMode As Boolean = False
    '=============================================================================================================
    'btn click refresh
    Public btnperformclick As Boolean = False
    '=============================================================================================================
    'add module
    'Auto Reset
    Public AutoInventoryReset As Boolean = False

    Public messageboxappearance As Boolean = False
    Public table As String
    Public fields As String
    Public value As String
    Public where As String
    Public successmessage As String
    Public errormessage As String
    Public returnvalrow As String
    Public mysqlcondition As String
    Public SystemLogType As String
    Public SystemLogDesc As String
    '=============================================================================================================
    'connection module
    Public myMySqlException As Boolean = True
    'Expenses
    Public hasbutton As Boolean = False
    '
    Public DatasourceOrAdd As Boolean = False
    'Payment Form
    Public TEXTBOXMONEYVALUE
    Public TEXTBOXTOTALPAYVALUE
    Public TEXTBOXCHANGEVALUE
    Public ifclose As Boolean
    'ModeOfTransactionDetails
    Public TEXTBOXFULLNAMEVALUE
    Public TEXTBOXREFERENCEVALUE
    '
    Public IfConnectionIsConfigured As Boolean
    'Loading
    Public ValidDatabaseLocalConnection As Boolean = False

    Public DISABLESERVEROTHERSPRODUCT As Boolean = False
    Public DisallowedCharacters As String = "!#$%'~`{}^¨|°¬+[]^¨\/,;=?<>*&()-+=" & """"

    Public SENIORDETAILSBOOL As Boolean = False
    Public SeniorDetailsID As String
    Public SeniorDetailsName As String

    Public ProductTotalPrice As Double = 0
    Public HASOTHERSLOCALPRODUCT As Boolean = False
    Public HASOTHERSSERVERPRODUCT As Boolean = False

    Public BegBalanceBool As Boolean = False

    Public HASUPDATE As Boolean = False

    Public LOGOUTFROMPOS As Boolean = True

    Public S_OLDGRANDTOTAL As Double = 0

End Module
