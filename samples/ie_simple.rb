require 'win32ole'

ie = WIN32OLE.new('InternetExplorer.Application')
puts "VISIBLE: #{ie.visible}"
puts "VISIBLE w/ []: #{ie['Visible']}"
puts "VISIBLE w/ Invoke: #{ie.invoke('Visible')}"
ie.Visible = TRUE  # Upper-case
puts "VISIBLE: #{ie.visible}"
sleep 1
ie['Visible'] = false
puts "VISIBLE: #{ie.visible}"
#puts ie.ole_methods
ie.gohome
puts "NAME: #{ie.name}"  # Lower-case
ie.quit 



