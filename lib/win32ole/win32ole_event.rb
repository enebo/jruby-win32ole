class WIN32OLE_EVENT
  def initialize(ole, event_name)
    @event_handlers = {}
    # TODO: Argument errors + specs

    if event_name.nil? # Default event name
      # TODO: get default event
    end

    org.jruby.ext.win32ole.RubyDispatchEvents.setupDispatchEventHandler(ole, self)
  end

  def on_event(name=nil, &block)
    if name
      @event_handlers[name.to_s] = block
    else
      @default_handler = block
    end
  end

  def method_missing(name, *args)
    name = name.to_s
    handler = @event_handlers[name]
    if handler
      handler.call *args
    elsif @default_handler
      @default_handler.call name, *args
    end
  end

  # Almost noop this.  We don't because it get CPU hot when people put this
  # in a hot loop!
  def self.message_loop
    DispatchEvents.message_loop
  end
end
