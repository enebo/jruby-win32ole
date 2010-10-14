require 'java'
require 'jacob.jar'

require 'win32ole/win32ole'      # <- java native impl of WIN32OLE
require 'win32ole/win32ole_ruby' # <- ruby impl of WIN32OLE

java_import java.util.Calendar

java_import com.jacob.com.Variant
java_import com.jacob.com.Automation
java_import com.jacob.com.ComFailException
java_import com.jacob.com.ComThread
java_import com.jacob.com.Dispatch
java_import com.jacob.com.DispatchEvents
java_import com.jacob.com.EnumVariant
java_import com.jacob.com.FuncDesc
java_import com.jacob.com.TypeInfo
java_import com.jacob.com.TypeLib
java_import com.jacob.com.VarDesc
java_import com.jacob.com.VariantUtilities

java_import org.jruby.ext.win32ole.RubyWIN32OLE
java_import org.jruby.ext.win32ole.RubyInvocationProxy


require 'win32ole/win32ole_error'
require 'win32ole/win32ole_method'
require 'win32ole/win32ole_variant'
require 'win32ole/win32ole_variable'
require 'win32ole/win32ole_event'
require 'win32ole/win32ole_param'
require 'win32ole/win32ole_type'
require 'win32ole/win32ole_typelib'
