VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindchillObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "__Modules"
Option Explicit
'@IgnoreModule ProcedureNotUsed

Private m_order_number As String
Private m_customer_release_OID As String
Private m_production_release_OID As String

'Declare windchill constants
Const M_URL_WINDCHILL_ODATA As String = "https://windchill.fcpm.biz/Windchill/servlet/odata/"
Const M_URL_WINDCHILL_DOCMGMT As String = M_URL_WINDCHILL_ODATA & "/v5/DocMgmt"
Const M_URL_WINDCHILL_BASE As String = "https://windchill.fcpm.biz:443/Windchill/servlet/rest"
Const M_URL_WINDCHILL_CSRF As String = M_URL_WINDCHILL_BASE & "/security/csrf"
Const M_URL_WINDCHILL_SEARCH As String = M_URL_WINDCHILL_BASE & "/search/objects?"
Const M_URL_WINDCHILL_DOCUMENT_DOWNLOAD As String = M_URL_WINDCHILL_BASE & "/files/downloadDoc/"

Const M_WINDCHILL_FILTER_STRING As String = "$filter=name eq "
Const M_WINDCHILL_TYPEID_STRING As String = "&typeId="

'===================================================
'                   START SUBS
'===================================================

'Get Windchill Data (OIDs and names)
Public Sub get_windchill_data(order_number_input As String, eng_type As String)

    Dim customer_release_name As String
    Dim production_release_name As String
    Dim windchill_search_term As String
    
    Dim typeID As String: typeID = "wt.doc.WTDocument"
    
    m_order_number = order_number_input
    
    If (eng_type = "PC") Then
        customer_release_name = m_order_number & "E"
        production_release_name = m_order_number & "ER"
    ElseIf (eng_type = "ME") Then
        customer_release_name = m_order_number & "M"
        production_release_name = m_order_number & "MR"
    End If
    
    windchill_search_term = M_WINDCHILL_FILTER_STRING & Chr$(39) & customer_release_name & Chr$(39) & M_WINDCHILL_TYPEID_STRING & typeID
    m_customer_release_OID = get_OID(windchill_search_term)
    windchill_search_term = M_WINDCHILL_FILTER_STRING & Chr$(39) & production_release_name & Chr$(39) & M_WINDCHILL_TYPEID_STRING & typeID
    m_production_release_OID = get_OID(windchill_search_term)
    
    
End Sub


'===================================================
'                   START FUNCTIONS
'===================================================
Private Function get_OID(search_term As String) As String

    Dim http_request As Object: Set http_request = CreateObject("MSXML2.XMLHTTP")
    Dim http_request_string As String
    Dim http_response_text As String
       
    'Parsed JSON Response Text
    Dim JSON As Dictionary

    'Search Returned
    Dim object_search_items As Dictionary
    
    http_request_string = M_URL_WINDCHILL_SEARCH & search_term
    
    With http_request
        .Open "GET", http_request_string, False
        .Send
    End With
    
    'TODO: Deal with no items found
    If http_request.Status <> 200 Then get_OID = vbNullString: Exit Function
    
    http_response_text = http_request.ResponseText
    
    Set JSON = JsonConverter.ParseJson(http_response_text)
    
    If (JSON.Item("items").Count > 0) Then
        Set object_search_items = JSON.Item("items").Item(1)
        get_OID = object_search_items.Item("id")
    Else
        get_OID = vbNullString
    End If
    
End Function

Property Get order_number() As String
    order_number = m_order_number
End Property

Property Get customer_release_OID() As String
    customer_release_OID = m_customer_release_OID
End Property

Property Get production_release_OID() As String
    production_release_OID = m_production_release_OID
End Property

Public Property Get URL_WINDCHILL_ODATA() As String
    URL_WINDCHILL_ODATA = M_URL_WINDCHILL_ODATA
End Property

Public Property Get URL_WINDCHILL_DOCMGMT() As String
    URL_WINDCHILL_DOCMGMT = M_URL_WINDCHILL_DOCMGMT
End Property

Public Property Get URL_WINDCHILL_BASE() As String
    URL_WINDCHILL_BASE = M_URL_WINDCHILL_BASE
End Property

Public Property Get URL_WINDCHILL_CSRF() As String
    URL_WINDCHILL_CSRF = M_URL_WINDCHILL_CSRF
End Property

Public Property Get URL_WINDCHILL_SEARCH() As String
    URL_WINDCHILL_SEARCH = M_URL_WINDCHILL_SEARCH
End Property

Public Property Get URL_WINDCHILL_DOCUMENT_DOWNLOAD() As String
    URL_WINDCHILL_DOCUMENT_DOWNLOAD = M_URL_WINDCHILL_DOCUMENT_DOWNLOAD
End Property

Public Property Get WINDCHILL_FILTER_STRING() As String
    WINDCHILL_FILTER_STRING = M_WINDCHILL_FILTER_STRING
End Property

Public Property Get WINDCHILL_TYPEID_STRING() As String
    WINDCHILL_TYPEID_STRING = M_WINDCHILL_TYPEID_STRING
End Property


'=====================================================
'       NOTE: Future use
'=====================================================
Private Sub get_CSRF_nonce()

    Dim http_request As Object: Set http_request = CreateObject("MSXML2.XMLHTTP")
    Dim http_request_string As String
    Dim http_response_text As String
    
    'Parsed JSON Response Text
    Dim JSON As Dictionary
    
    http_request_string = M_URL_WINDCHILL_CSRF
    
    With http_request
        .Open "GET", http_request_string, False
        .Send
    End With
    
    http_response_text = http_request.ResponseText
    Set JSON = JsonConverter.ParseJson(http_response_text)

    Dim itemsDictionary As Dictionary
    Dim attributesDictionary As Dictionary
    
    Dim csrfNonce As String
    If (JSON.Item("items").Count > 0) Then
        Set itemsDictionary = JSON.Item("items").Item(1)
        Set attributesDictionary = itemsDictionary.Item("attributes")
        csrfNonce = attributesDictionary.Item("nonce")
    End If

End Sub

