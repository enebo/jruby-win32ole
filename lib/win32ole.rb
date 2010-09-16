require 'java'
require 'jacob.jar'

require 'win32ole/win32ole'      # <- java native impl of WIN32OLE
require 'win32ole/win32ole_ruby' # <- ruby impl of WIN32OLE

import java.util.Calendar

import com.jacob.com.Variant
import com.jacob.com.Automation
import com.jacob.com.ComFailException
import com.jacob.com.ComThread
import com.jacob.com.Dispatch
import com.jacob.com.EnumVariant
import com.jacob.com.TypeInfo
import com.jacob.com.TypeLib
import com.jacob.com.VarDesc

import com.jacob.com.VariantUtilities

require 'win32ole/win32ole_error'
require 'win32ole/win32ole_method'
require 'win32ole/win32ole_variant'
require 'win32ole/win32ole_variable'
require 'win32ole/win32ole_event'
require 'win32ole/win32ole_type'
require 'win32ole/win32ole_typelib'
