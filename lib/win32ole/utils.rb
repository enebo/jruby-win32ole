class WIN32OLE
  module Utils
    def from_variant(value)
      object = VariantUtilities.variant_to_object(value)
      object.kind_of?(Dispatch) ? WIN32OLE.new(object) : object
    end
  end
end
