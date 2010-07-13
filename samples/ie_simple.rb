require 'win32ole'

ie = WIN32OLE.new('InternetExplorer.Application')
puts "VISIBLE: #{ie.visible}"
ie.visible = TRUE
puts "VISIBLE: #{ie.visible}"
ie.gohome

