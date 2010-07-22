class WIN32OLE_TYPE
  def initialize(typelib, info, docs)
    @typelib, @info, @docs = typelib, info, docs
  end

  def name
    @docs.name
  end

  def major_version
    @info.major_version
  end

  def minor_version
    @info.minor_version
  end

  def typekind
    @info.typekind
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
end
