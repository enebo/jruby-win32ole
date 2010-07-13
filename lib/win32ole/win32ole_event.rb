
class RubyInvocationProxy < com.jacob.com.InvocationProxy
  def initialize(target)
    super()
    @target = target
  end

  def invoke(name, *parameters)
    puts "In invoke"
    @target.__send__ name, *parameters
  end
end

class RubyDispatchEvents < com.jacob.com.DispatchEvents
  def initialize(source, event_sink)
    super(source, event_sink)
  end

  def getInvocationProxy(target)
    puts "HERE! yes #{target}"
    RubyInvocationProxy.new(target)
  end
end

class WIN32OLE_EVENT
  def initialize(ole, event_name)
    # TODO: Argument errors + specs

    if event_name.nil? # Default event name
      # TODO: get default event
    end
    
    RubyDispatchEvents.new(ole.dispatch, self)
  end

  def method_missing(name, *args)
    puts "Called #{name} #{args.join(',' )}"
  end
end
