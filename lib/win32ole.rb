require 'java'
require 'jacob.jar'

import com.jacob.com.ComThread
import com.jacob.com.Dispatch
import com.jacob.com.EnumVariant
import com.jacob.com.Variant
import com.jacob.com.VariantUtilities

require 'win32ole/user32' # FFI bindings to user32 lib for message_loop funs
require 'win32ole/win32ole'
require 'win32ole/win32ole_variant'
require 'win32ole/win32ole_event'
