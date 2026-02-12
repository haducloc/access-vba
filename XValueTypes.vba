Option Compare Database
Option Explicit

' Value Types
Public Enum XValueType
    Type_String = 1
    Type_Bool = 2
    
    Type_Byte = 3
    Type_UByte = 4

    Type_Int2 = 5
    Type_Int4 = 6
    Type_Int8 = 7

    Type_Float = 8
    Type_Double = 9
    Type_Decimal = 10

    Type_Date = 11
    Type_Time = 12
    Type_DateTime = 13

    Type_GUID = 14
End Enum
