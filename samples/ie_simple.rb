require 'win32ole'

ie = WIN32OLE.new('InternetExplorer.Application')
puts "VISIBLE: #{ie.visible}"
ie.Visible = TRUE  # Upper-case
puts "VISIBLE: #{ie.visible}"
ie.gohome
puts "NAME: #{ie.name}"  # Lower-case
ie.quit 


