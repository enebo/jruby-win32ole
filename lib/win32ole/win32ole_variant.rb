class WIN32OLE
  module VARIANT
    VT_I2 = 2 # Short
    VT_I4 = 3 # Int
    VT_R4 = 4 # Float
    VT_R8 = 5 # Double
    VT_CY = 6 # Currency
    VT_DATE = 7 # Date
    VT_BSTR = 8 # String
    VT_DISPATCH = 9 # Dispatch
    VT_ERROR = 10 # Error
    VT_BOOL = 11 # Boolean
    VT_VARIANT = 12 # Variant containing Variant
    VT_UNKNOWN = 13 # Unknown
    VT_I1 = 16 # Nothing in Jacob
    VT_UI1 = 17 # Byte
    VT_UI2 = 18 # Nothing in Jacob
    VT_UI4 = 19 # Nothing in Jacob
    VT_I8 = 20 # Not in MRI win32ole but in Jacob
    VT_UI8 = 21 # !Jacob
    VT_INT = 22 # Nothing in Jacob
    VT_UINT = 23 # Nothing in Jacob
    VT_VOID = 24 # !Jacob
    VT_HRESULT = 25 # !Jacob
    VT_PTR = 26 # Pointer
    VT_SAFEARRAY = 27 # !Jacob
    VT_CARRAY = 28 # !Jacob
    VT_USERDEFINED = 29 # !Jacob
    VT_ARRAY = 8192 # Array
    VT_BYREF = 16384 # Reference
  end
end
