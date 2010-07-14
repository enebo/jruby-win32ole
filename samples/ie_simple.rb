require 'win32ole'

ie = WIN32OLE.new('InternetExplorer.Application')
puts "VISIBLE: #{ie.visible}"
puts "VISIBLE w/ []: #{ie['Visible']}"
ie.Visible = TRUE  # Upper-case
puts "VISIBLE: #{ie.visible}"
sleep 1
ie['Visible'] = false
puts "VISIBLE: #{ie.visible}"
ie.gohome
puts "NAME: #{ie.name}"  # Lower-case
ie.quit 


