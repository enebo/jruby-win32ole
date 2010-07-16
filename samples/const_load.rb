require 'win32ole'

module FOO
end

ie = WIN32OLE.new('InternetExplorer.Application')
WIN32OLE.const_load(ie, FOO)
puts FOO.constants
