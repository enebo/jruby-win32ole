class WIN32OLE_VARIABLE
  attr_reader :name

  def initialize(type, var_desc, name)
    @type, @var_desc, @name = type, var_desc, name
  end

  # TODO: Implement ole_type, ole_type_detail

  def value
    from_variant(@var_desc.constant)
  end

  def variable_kind
    variable_kind_string(varkind)
  end

  def varkind
    @var_desc.varkind
  end

  def inspect
    "#<WIN32OLE_VARIABLE:#{to_s}=#{value.inspect}>"
  end

  alias :to_s :name

  def visible?
    flags = @var_desc.flags
    flags & (VarDesc::VARFLAG_FHIDDEN | VarDesc::VARFLAG_FRESTRICTED | VarDesc::VARFLAG_FNONBROWSABLE) == 0
  end

  include WIN32OLE::Utils
end
