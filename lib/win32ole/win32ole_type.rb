class WIN32OLE_TYPE
  attr_reader :typeinfo

  def initialize(*args)
    case args.length
    when 2 then 
      typelib_name, olename = SafeStringValue(args[0]),SafeStringValue(args[1])
      @typelib = WIN32OLE_TYPELIB.new(typelib_name) # Internal call
      find_all_typeinfo(@typelib.typelib) do |info, docs|
        if (docs.name == olename)
          @typeinfo, @docs = info, docs
          break
        end
      end
    when 3 then
      @typelib, @typeinfo, @docs = *args
    else
      raise ArgumentError.new("wrong number of arguments (#{args.length} for 2)")
    end
  end

  def guid
    @typeinfo.guid
  end

  def helpcontext
    @docs.help_context
  end

  def helpstring
    @docs.doc_string
  end

  def helpfile
    @docs.help_file
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
    all_methods(@typeinfo) do |ti, oti, desc, docs, index|
      members << WIN32OLE_METHOD.new(self, ti, oti, desc, docs, index)
      nil
    end
    members
  end

  def variables
    variables = []
    all_vars(@typeinfo) do |desc, name|
      variables << WIN32OLE_VARIABLE.new(self, desc, name)
    end
    variables
  end

  def visible?
    @typeinfo.flags & (TypeInfo::TYPEFLAG_FHIDDEN | TypeInfo::TYPEFLAG_FRESTRICTED) == 0
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
