class WIN32OLE_EVENT
  def initialize(ole, event_name)
    # TODO: Argument errors + specs

    if event_name.nil? # Default event name
      # TODO: get default event
    else
      id = Dispatch.getIDOfName(ole, name)
    end
    
  end
end
