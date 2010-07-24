class WIN32OLE_TYPE
  attr_reader :typeinfo

  def initialize(*args)
    case args.length
    when 2 then 
      typelib_name, olename = *args
      @typelib = WIN32OLE_TYPELIB.new(typelib_name) # Internal call
      find_all_typeinfo(@typelib.typelib) do |info, docs|
        if (docs.name == olename)
          @typeinfo = info
          break
        end
      end
    when 3 then
      @typelib, @typeinfo, @docs = *args
    else
      raise ArgumentError.new("wrong number of arguments (#{args.length} for 2)")
    end
  end

  def name
    @docs.name
  end

  def major_version
    @typeinfo.major_version
  end

  def minor_version
    @typeinfo.minor_version
  end

  def ole_methods
    members = []
    all_methods(@typeinfo) do |ti, oti, desc, docs, name|
      members << WIN32OLE_METHOD.new(nil, name, ti, oti, desc.memid)
      nil
    end
    members
  end

  def typekind
    @typeinfo.typekind
  end

  class << self
    # This is obsolete, but easy to emulate
    def typelibs
      WIN32OLE_TYPELIB.typelibs.collect {|t| t.name }
    end

    def ole_classes(tlib)
      WIN32OLE_TYPELIB.ole_classes(tlib)
    end
  end

  include WIN32OLE::Utils
end
